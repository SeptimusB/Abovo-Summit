Imports System.ComponentModel
Imports System.Data
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.AbovoExtendedDEControls
Imports Abovo.AbovoRangeDataSource
Imports Abovo.AbovoUnboundSource
Imports Abovo.ChangeLogManager
Imports Abovo.CustomGrid
Imports Abovo.DataObject
Imports Abovo.DefaultHelpers
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports Abovo.LogDebugDev
Imports Abovo.MasterChangeLog
Imports Abovo.PresentationManager
Imports Abovo.RepositaryItems
Imports DataInterfaceTemplate.AbovoGridRespoitaryCombo
Imports DevExpress
Imports DevExpress.Accessibility
Imports DevExpress.CodeParser
Imports DevExpress.Data
Imports DevExpress.Pdf.Native
Imports DevExpress.Pdf.Native.BouncyCastle.Asn1.X509.Qualified
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.Utils.Behaviors.Common
Imports DevExpress.Utils.Drawing
Imports DevExpress.Utils.Drawing.Helpers
Imports DevExpress.Utils.Extensions
Imports DevExpress.Utils.Gesture
Imports DevExpress.Utils.Layout
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraCharts
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraExport.Helpers
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Tab
Imports DevExpress.XtraGrid.Views
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraLayout
Imports DevExpress.XtraLayout.Customization
Imports DevExpress.XtraLayout.Utils
Imports DevExpress.XtraPrinting.Export
Imports DevExpress.XtraPrinting.Native
Imports DevExpress.XtraRichEdit.Commands
Imports DevExpress.XtraRichEdit.Import.Rtf
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraSpreadsheet
Imports DevExpress.XtraSpreadsheet.Commands.Internal
Imports DevExpress.XtraSpreadsheet.Model
Imports DevExpress.XtraTab
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Rows
Imports Microsoft.Win32
Public Class DataInterfaceTemplate

    Inherits DevExpress.XtraEditors.XtraForm



    'Public CollectionVariables
    Public ReliantNRs As List(Of String)
    Public AmActivated As Boolean = True
    Public DITName As String
    'Higher Level Classes/Objects
    Private DataPM As PresentationManager
    Private ActionMan As ActionManager
    Private DataPres As PresentationManager.DataPresentation
    Private Formatter As ObjectFormatter
    Private rs As New Resizer
    'Private TTController As ToolTipController
    Private ParentGroupForm As GroupInterfaceTemplate
    Private ActiveSpreadsheet As DevExpress.Spreadsheet.Worksheet
    Private ActiveWorkbook As IWorkbook
    'InterfaceTag
    Private InterfaceTag As AbovoInterfaceTag
    Private CalcEngID As Integer

    'Private Data Variables
    Private ModelID As Integer
    Private PresID As Integer
    Private GSID As Integer
    Private CSID As Integer
    Private CustCalcEditInteger As RepositoryItemCalcEdit
    Private CustCalcEditDecimal As RepositoryItemCalcEdit
    Private UnboundDataSources() As AbovoUnboundSource
    Private RangeDataSources() As RangeDataSource
    Private UBSDataSourceCount As Integer = -1
    Private RangeDataSourceCount As Integer = -1
    Private DblClickCell As Boolean = False
    Private MyData As DataObject
    Private PresentedDS As Abovo.DataObject.DataCellRange
    Private PresentedColumn As Abovo.DataObject.SheetDataColumn
    Private ChangeMan As ModelChangeManager

    'Layout Variables
    Private ScaleUnits As Single
    Private Scalefactor As Single
    Private IdealGridRowHeight As Integer = 54
    Private MinGridRowHeight As Integer = 20

    'State Variables
    Private BIsDirty As Boolean

    Private ControlsInitialised As Boolean = False
    Private InterfaceResourcesReleased As Boolean = False
    Private FooterOn As Boolean
    Private FooterDone As Boolean
    Private DIInitialised As Boolean = False
    Private SheetIInitialised As Boolean = False
    Private InterfaceMode As String
    Private RightClick As Boolean = False

    'Placeholder variables for constructors
    Private PropertiesCount As Integer
    Private PropertyList As IEnumerable(Of UnboundSourceProperty)
    Private PropType As System.Type
    Private ColList As List(Of String)
    Private TPs() As TablePanel
    Private XtraTabPages() As XtraTabPage
    Private XtraTabPageCount As Integer = 0
    Private ColCount As Integer = -1
    Private ColName As String
    'Private ColExtenders() As ColumnButtonExtender
    'Private ColExtenderCount As Integer = -1
    Private BandExtenders() As BandButtonExtender
    Private VGridCategoryExtenders As New List(Of VGridCategoryButtonExtender)
    Private VGridCategoryExtendersBySection As New Dictionary(Of Integer, List(Of VGridCategoryButtonExtender))
    Private VGridInplaceEditorHelpersBySection As New Dictionary(Of Integer, List(Of VGridRowInplaceEditorHelper))

    Private Class VGridLayoutTag

        Public TableRowIndex As Integer = -1
        Public IsLiveGrid As Boolean = False

    End Class

    Private Class LiveVGridRowTag

        Public SourceColumnIndex As Integer

    End Class

    Private Class LiveVGridCategoryTag

        Public SourceRowIndex As Integer
        Public SourceColumnIndex As Integer

    End Class

    'A MultiEditorRow represents more than one DataColumnTag.  Keep the
    'constituent tags together so VGrid handlers can resolve the correct
    'underlying spreadsheet/data column by CellIndex.
    Private Class VGridMultiEditorRowTag

        Public ColumnTags() As DataColumnTag

        Public Sub New(ParamArray ByVal Tags() As DataColumnTag)
            ColumnTags = Tags
        End Sub

        Public Function GetColumnTag(ByVal CellIndex As Integer) As DataColumnTag

            If ColumnTags Is Nothing OrElse ColumnTags.Length = 0 Then Return Nothing

            If CellIndex < 0 OrElse CellIndex >= ColumnTags.Length Then
                CellIndex = 0
            End If

            Return ColumnTags(CellIndex)

        End Function

    End Class

    Private Class JointVentureVGridRowTag
        Public CategoryName As String
        Public RowKind As String
        Public ValueDataFormat As String
    End Class

    Private Class JointVentureCellBinding
        Public SourceSheet As String
        Public SourceAddress As String
        Public DataFormat As String
        Public IsReadOnly As Boolean

        'Equivalent of DataColumnTag.HasRules for the dedicated JV renderer.
        'The generic Funding VGrid receives this from Structure.xml through
        'DataManager; JV is built directly from workbook ranges, so retain the
        'same metadata on the per-cell binding.
        Public HasRules As Boolean
    End Class

    'VGrid record separators are painted outside the individual
    'CustomDrawRowValueCell operation.  For the handful of worksheet-style
    'areas where a separator is visually misleading, cache the separator
    'segments during cell custom-draw and erase them in the control's final
    'Paint event.
    Private Class JointVentureSeparatorSuppression
        Public X As Integer
        Public Top As Integer
        Public Bottom As Integer
        Public BackColor As Color
    End Class

    Private JointVentureSeparatorSuppressions As New Dictionary(
        Of VGridControl,
        Dictionary(Of String, JointVentureSeparatorSuppression))

    Private Sub QueueJointVentureSeparatorSuppression(
        ByVal Grid As VGridControl,
        ByVal Key As String,
        ByVal X As Integer,
        ByVal Top As Integer,
        ByVal Bottom As Integer,
        ByVal BackColor As Color)

        If Grid Is Nothing Then Return

        Dim GridSuppressions As Dictionary(
            Of String,
            JointVentureSeparatorSuppression) = Nothing

        If Not JointVentureSeparatorSuppressions.TryGetValue(
            Grid,
            GridSuppressions) Then

            GridSuppressions =
                New Dictionary(
                    Of String,
                    JointVentureSeparatorSuppression)(
                        StringComparer.Ordinal)

            JointVentureSeparatorSuppressions(Grid) =
                GridSuppressions

        End If

        GridSuppressions(Key) =
            New JointVentureSeparatorSuppression With {
                .X = X,
                .Top = Top,
                .Bottom = Bottom,
                .BackColor = BackColor
            }

    End Sub

    Private Sub JointVenture_PostPaint(ByVal sender As Object,
                                       ByVal e As PaintEventArgs)

        Dim Grid As VGridControl =
            TryCast(sender, VGridControl)

        If Grid Is Nothing Then Return

        Dim GridSuppressions As Dictionary(
            Of String,
            JointVentureSeparatorSuppression) = Nothing

        If Not JointVentureSeparatorSuppressions.TryGetValue(
            Grid,
            GridSuppressions) Then Return

        If GridSuppressions.Count = 0 Then Return

        'This Paint event occurs after the VGrid's normal element painting.
        'Erase only the queued vertical separator segments.  Do not disable
        'VGrid vertical lines globally because normal data cells still need them.
        For Each Suppression As JointVentureSeparatorSuppression In
            GridSuppressions.Values

            If Suppression.Bottom <= Suppression.Top Then Continue For

            Using ErasePen As New Pen(
                Suppression.BackColor,
                2)

                e.Graphics.DrawLine(
                    ErasePen,
                    Suppression.X,
                    Suppression.Top,
                    Suppression.X,
                    Suppression.Bottom)

            End Using

        Next

        'The rectangles are refreshed by CustomDrawRowValueCell on every paint
        'cycle. Clearing here prevents stale coordinates after horizontal
        'scrolling, resizing or BestFit/layout changes.
        GridSuppressions.Clear()

    End Sub

    Private JointVentureTables As New Dictionary(Of VGridControl, DataTable)
    Private JointVentureBindings As New Dictionary(Of VGridControl, Dictionary(Of String, JointVentureCellBinding))
    Private JointVentureTextEditors As New Dictionary(Of VGridControl, RepositoryItemTextEdit)
    Private JointVentureYearEditors As New Dictionary(Of VGridControl, RepositoryItemComboBox)
    Private JointVentureMoneyEditors As New Dictionary(Of VGridControl, RepositoryItemTextEdit)
    Private JointVenturePercentEditors As New Dictionary(Of VGridControl, RepositoryItemSpinEdit)

    Private Class InterfaceSectionRuntimeState
        Public SectionID As Integer

        'IsBuilt describes the WinForms/DevExpress control tree only.
        Public IsBuilt As Boolean = False

        'IsDirty means the section needs rebuilding before it is shown.
        Public IsDirty As Boolean = True

        'This is deliberately separate from IsBuilt.
        '
        'A worksheet structural change can invalidate a section that has never
        'been lazy-built, or one whose controls have already been unloaded.
        'In either case IsBuilt=False, but the PresentationSection/DataCellRange
        'must still be regenerated because its CellDataPoint.SourceAddress map
        'may have changed.
        Public NeedsPresentationRedefinition As Boolean = False

        Public SourceWorksheets As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    End Class

    Private SectionRuntimeStates As New Dictionary(Of Integer, InterfaceSectionRuntimeState)
    Private SuppressLazyTabEvents As Boolean = False
    Private SkipNextRepositoryRefreshSection As Integer = -1

    'When a lazy page is first selected, hold one outer XtraTabControl update
    'across BOTH SelectedPageChanging and SelectedPageChanged. BuildSection has
    'its own nested BeginUpdate/EndUpdate, but this outer lock prevents the old
    'page from repainting between the off-screen build and the final selected-
    'page sizing pass.
    Private LazyTabTransitionUpdateHeld As Boolean = False

    'BeginUpdate suppresses DevExpress updates, but complex lazy builds can still
    'cause native child windows on the currently selected page to repaint after
    'their layout changes. Freeze the complete form at the Win32 paint layer for
    'the very short first-time lazy transition.
    Private LazyTabTransitionRedrawDisabled As Boolean = False

    'Exactly one application-level mouse-wheel filter belongs to this interface.
    'It dynamically routes against XtraTabControlNewGIT.SelectedTabPage.
    '
    'Historically a new IMessageFilter was added for EVERY tab page and the
    'instances were not retained/removed. Because message filters are
    'Application-wide, repeated interface construction could accumulate filters
    'and forward the same wheel message more than once.
    Private InterfaceScrollRediverter As ScrollRediverter

    'Selected VGrid sections are never destroyed synchronously from the action /
    'dependency notification path.  A short WinForms timer gives DevExpress a
    'complete message-loop boundary after MouseUp/modal input processing before
    'the selected section is disposed and rebuilt.
    Private PendingSectionRebuilds As New HashSet(Of Integer)
    Private SectionRebuildTimer As System.Windows.Forms.Timer
    Private InHeaderCombos() As AbovoRespositaryItem
    Private InHeaderCombosCount As Integer = -1
    Private BandExtenderCount As Integer = -1
    Private FixedLeftIndent = 200

    'Static Controls
    Private AcControl As AccordionControl

    'Dynamic Arrays
    Private PropertyArray() As UnboundSourceProperty

    'Arrays for Dynamic Contols
    Private GridControls() As GridControl
    Private VertGridControls() As VGridControl
    Private ExpanableGridControls() As GridControl
    Private ExpanableGridControlsCount As Integer = -1
    Private TextBoxes() As AbovoDETextEdit
    Private DateBoxes() As AbovoDEDateEdit
    Private Labels() As DevExpress.XtraEditors.LabelControl
    Private SpinEdits() As AbovoDESpinEdit
    Private Combos() As DevExpress.XtraEditors.ComboBoxEdit
    Private AcElements() As AccordionControlElement
    Private AcElementlist As List(Of AccordionControlElement)
    Private UsedGridVIEWS(-1) As GridView
    Private UsedBANDedGridVIEWS(-1) As BandedGridView
    Private UsedBANDedGridViewBANDS(-1) As GridBand
    Private UsedBANDedGridViewCOLS(-1) As BandedGridColumn
    Private AcContainers() As AccordionContentContainer
    Private hyperlinkLabelControls() As HyperlinkLabelControl
    Private DontAddCDHDef As Boolean = False
    Private RefreshableControls() As Object
    Private HaveCalledData As Boolean = False
    Private ComboReposClasses() As AbovoGridRespoitaryCombo
    Private AbovoTabPages() As AbovoTabPage
    Private AbovoTabPagesCount As Integer = -1
    Private CurrentAbovoTabPage As AbovoTabPage

    'Layout Control variables



    'MappedTable Variables

    Private ADELabels() As AbovoDELabel
    Private ADELabelsCount As Integer = -1
    Private ADEDateEdits() As AbovoDEDateEdit
    Private ADEDateEditsCount As Integer = -1
    Private ADEHyperlinks() As AbovoDEHyperlinkLabel
    Private ADEHyperlinksCount As Integer = -1
    Private ADETextEdits() As AbovoDETextEdit
    Private ADETextEditsCount As Integer = -1
    Private ADEComboBoxes() As AbovoDEComboBox
    Private ADEComboBoxesCount As Integer = -1
    Private ComboReposClassesCount As Integer = -1
    'Initialisation Integers for Arrays
    Private RefreshableControlsCount As Integer = -1
    Private TextEditCount As Integer = -1
    Private BandExtendersCount As Integer = -1
    Private DateEditCount As Integer = -1
    Private CombosCount As Integer = -1
    Private GridCombosCount As Integer = -1
    Private SpinEditCount As Integer = -1
    Private ComboEditCount As Integer = -1
    Private LabelsCount As Integer = -1
    Private GridCount As Integer = -1
    Private VertGridCount As Integer = -1
    Private AcControlCount As Integer = -1
    Private AcElementCount As Integer = -1
    Private GridViewCount As Integer = -1
    Private AcContainersCount As Integer = -1
    Private BandGridViewsCount As Integer = -1
    Private BandGridViewBandsCount As Integer = -1
    Private BandGridViewColsCount As Integer = -1

    Private ActiveLinkElement As ElementInterfaceLinkTag

    'Interface controls
    Private WindowsUIButtonPanelSaveClose As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()

    'Development Variables
    Private DataCallCount As Integer = 0


    Public Sub New(SetModelID As Integer, Optional ByVal SetGSID As Integer = -1, Optional ByVal SetCSID As Integer = -1, Optional ByVal MyParent As GroupInterfaceTemplate = Nothing, Optional ByVal SetInterfaceMode As String = "Normal", Optional ByVal Interfacelink As ElementInterfaceLinkTag = Nothing)

        ParentGroupForm = MyParent

        InitializeComponent()

        If Not IsNothing(MyParent) Then

            'The interface previously followed only the parent's WIDTH.
            'Its HEIGHT therefore remained at the form's design-time/default
            'height, which in turn capped XtraTabControlNewGIT / XtraTabPage.
            '
            'Use the parent's CLIENT area for both dimensions.
            Me.Size =
                New Size(
                    Math.Max(1, MyParent.ClientSize.Width),
                    Math.Max(1, MyParent.ClientSize.Height))

        End If

        ActiveWorkbook = ExcelModels(SetModelID).WB

        ActiveLinkElement = Interfacelink

        AddMenuInterface(InterfaceMode)

        Formatter = New ObjectFormatter

        InterfaceMode = SetInterfaceMode
        GSID = SetGSID
        CSID = SetCSID
        ModelID = SetModelID
        FooterOn = False

        DITName = GetCSName(SetModelID, SetGSID, SetCSID)

        ChangeMan = ExcelModels(SetModelID).ChangeManager
        DataPM = ExcelModels(SetModelID).WBDataPres
        Me.ForeColor = ExcelModels(SetModelID).ColourSwatch

        Dim CheckTrans As AbovoTransaction
        CheckTrans = DataPM.ValidatePresentation(GSID, CSID)

        If CheckTrans.BError = False Then

            PresID = CheckTrans.IntegerReturn
            DataPres = DataPM.DataPresentations(PresID)

        End If

        CalcEngID = ExcelModels(ModelID).WBCalcEngine.AddActiveObject(Me)

        If DataPres.DefaultWorksheet IsNot Nothing Then

            ActiveSpreadsheet = ExcelModels(SetModelID).WB.Worksheets(DataPres.DefaultWorksheet)

            ExcelModels(ModelID).WBCalcEngine.AddActiveWorksheet(CalcEngID, ActiveSpreadsheet)

        End If

        'XtraTabControlDI = New XtraTabControl With {
        '    .Dock = DockStyle.Fill
        '    }

        'TablePanelDIT.Controls.Add(XtraTabControlDI)
        'TablePanelDIT.SetCell(XtraTabControlDI, 1, 0)
        'TTController = New ToolTipController With {
        '    .ToolTipType = ToolTipType.Default,
        '    .Rounded = True}

        'Create the shared CalcEdit repository items before any grids are built.
        'PopulateDataInterface uses these repository items.
        CustCalcEditInteger = New RepositoryItemCalcEdit With {
            .TextEditStyle = TextEditStyles.Standard,
            .UseMaskAsDisplayFormat = True,
            .UseAdvancedMode = DefaultBoolean.True,
            .Precision = 0
        }

        CustCalcEditDecimal = New RepositoryItemCalcEdit With {
            .TextEditStyle = TextEditStyles.Standard,
            .UseMaskAsDisplayFormat = True,
            .UseAdvancedMode = DefaultBoolean.True,
            .Precision = 5
        }

        AddHandler CustCalcEditInteger.Click, AddressOf CalcEditPopup
        AddHandler CustCalcEditDecimal.Click, AddressOf CalcEditPopup

        If WorkMode = "INTERFACE" Then


            PopulateDataInterface()

        ElseIf WorkMode = "SPREADSHEET" Then


            PopulateSpreadhsheetInterface()

        End If

        If ActiveLinkElement IsNot Nothing Then ProcessLinkElement()

        ControlsInitialised = True


        ResizeFonts()


        Exit Sub

    End Sub
    'This is added to the recently created CalcEdit for single click show
    Sub CalcEditPopup(Sender As Object, e As EventArgs)

        Dim a As CalcEdit = Sender
        'a.ShowPopup()

    End Sub

#Region "Public Methods / Properties"

    Public Shadows Sub Deactivate()

        ExcelModels(ModelID).WBCalcEngine.RemoveActiveObject(CalcEngID)

        AmActivated = False

    End Sub
    Public Sub Reactivate()

        ExcelModels(ModelID).WBCalcEngine.AddActiveObject(Me)

        ExcelModels(ModelID).WBCalcEngine.CalcFile()

        AmActivated = True

        EnsureSectionBuilt(XtraTabControlNewGIT.SelectedTabPageIndex)
        RefreshData()
        UpdateTabPage()

    End Sub

    Public Sub RefreshData()

        If Not IsNothing(DataPres) AndAlso GridControls IsNot Nothing Then

            For Each gridControl In GridControls

                If Not IsNothing(gridControl) AndAlso Not gridControl.IsDisposed Then

                    gridControl.RefreshDataSource()

                    'LiveGrid headings are workbook formula results too. Refresh
                    'their captions in place after calculation so a renamed stock
                    'type is reflected without rebuilding the whole interface.
                    RefreshLiveGridHeaders(gridControl)

                    'Data may have changed substantially since the grid was first
                    'built (paste, structural insert, workbook recalculation, etc).
                    'Grow columns/control width when the refreshed content now
                    'requires more space. Existing user-expanded widths are not
                    'shrunk by this refresh-time pass.
                    AutoFitGridAfterDataRefresh(gridControl)

                End If

            Next

        End If

        If VertGridControls IsNot Nothing Then

            For Each vertGridControl As VGridControl In VertGridControls

                If vertGridControl IsNot Nothing AndAlso Not vertGridControl.IsDisposed Then
                    vertGridControl.RefreshDataSource()
                    RefreshLiveVGridHeaders(vertGridControl)
                End If

            Next

        End If

SkipGridRefresh:

        If RefreshableControls Is Nothing Then GoTo SkipRefresh

        For Each RefControl As Object In RefreshableControls
            Try
                If Not IsNothing(RefControl) Then
                    RefControl.RefreshData()
                End If
            Catch ex As Exception
                ' Handle exception
                MsgBox("Error refreshing control: " & ex.Message)
            End Try
        Next

SkipRefresh:

        UpdateAllRules()

    End Sub



#End Region

#Region "Interface Methods"
    Public Sub RebuildInterface()

        If WorkMode = "INTERFACE" Then

            ClearDataInterface()

            PopulateDataInterface()

        ElseIf WorkMode = "SPREADSHEET" Then

            PopulateSpreadhsheetInterface()

        End If

    End Sub
    Public Sub ClearDataInterface()

        If SectionRebuildTimer IsNot Nothing Then
            SectionRebuildTimer.Stop()
            RemoveHandler SectionRebuildTimer.Tick, AddressOf ProcessQueuedSectionRebuilds
            SectionRebuildTimer.Dispose()
            SectionRebuildTimer = Nothing
        End If
        PendingSectionRebuilds.Clear()

        If ExcelModels IsNot Nothing AndAlso
           ModelID >= 0 AndAlso
           ModelID < ExcelModels.Length AndAlso
           ExcelModels(ModelID) IsNot Nothing AndAlso
           ExcelModels(ModelID).InterfaceDependencies IsNot Nothing Then

            ExcelModels(ModelID).InterfaceDependencies.UnregisterInterface(Me)

        End If

        If VGridCategoryExtenders IsNot Nothing Then

            For Each Extender As VGridCategoryButtonExtender In VGridCategoryExtenders

                If Extender IsNot Nothing Then

                    Try
                        Extender.DetachForDisposal()
                    Catch
                        'The owning VGrid may already have been disposed.
                    End Try

                End If

            Next

            VGridCategoryExtenders.Clear()

        End If

        If Not IsNothing(XtraTabControlDI) Then

            XtraTabControlDI.TabPages.Clear()

            GridControls = Nothing
            VertGridControls = Nothing
            UsedGridVIEWS = Nothing
            'UnboundDataSources = Nothing
            RangeDataSources = Nothing
            UBSDataSourceCount = -1
            RangeDataSourceCount = -1
            GridCount = -1
            VertGridCount = -1
            GridViewCount = -1

        End If

        ControlsInitialised = False

    End Sub

    Friend Sub ReleaseInterfaceResources()

        If InterfaceResourcesReleased Then Return
        InterfaceResourcesReleased = True
        AmActivated = False

        Try
            ResumeLazyTabTransitionRedraw()
            EndLazyTabTransitionUpdate()
            RemoveInterfaceScrollRediverter()
        Catch
            'Continue releasing data bindings and controls.
        End Try

        Try
            If SectionRebuildTimer IsNot Nothing Then
                SectionRebuildTimer.Stop()
                RemoveHandler SectionRebuildTimer.Tick, AddressOf ProcessQueuedSectionRebuilds
                SectionRebuildTimer.Dispose()
                SectionRebuildTimer = Nothing
            End If
        Catch
            SectionRebuildTimer = Nothing
        End Try

        PendingSectionRebuilds.Clear()

        If ExcelModels IsNot Nothing AndAlso
           ModelID >= 0 AndAlso ModelID < ExcelModels.Length AndAlso
           ExcelModels(ModelID) IsNot Nothing Then

            Dim model As ExcelModel = ExcelModels(ModelID)

            Try
                If model.InterfaceDependencies IsNot Nothing Then
                    model.InterfaceDependencies.UnregisterInterface(Me)
                End If
            Catch
                'Continue with control/data-source disposal.
            End Try

            Try
                If model.WBCalcEngine IsNot Nothing Then
                    model.WBCalcEngine.RemoveActiveObject(CalcEngID)
                End If
            Catch
                'Continue with control/data-source disposal.
            End Try
        End If

        'Detach data callbacks before controls, views, or workbook services are
        'disposed. This prevents DevExpress requesting data from a closing model.
        ClearAllGrids()

        If RangeDataSources IsNot Nothing Then
            For Each source As RangeDataSource In RangeDataSources
                If source Is Nothing Then Continue For

                Try
                    source.Dispose()
                Catch
                    'Other bindings and controls must still be released.
                End Try
            Next
        End If

        RangeDataSources = Nothing
        RangeDataSourceCount = -1

        ClearAllTPs()
        DataPres = Nothing
        ControlsInitialised = False

    End Sub
    Sub PopulateSpreadhsheetInterface()

    End Sub
    Private Sub PopulateDataInterface()

        ResetTimer(Me.Name & " Populate Data Interface")

        If DataPres Is Nothing OrElse DataPres.Sections Is Nothing OrElse DataPres.Sections.Length = 0 Then Exit Sub

        WindowsFormsSettings.SmartMouseWheelProcessing = True

        AcElementlist = New List(Of AccordionControlElement)
        SectionRuntimeStates.Clear()
        VGridCategoryExtendersBySection.Clear()
        VGridInplaceEditorHelpersBySection.Clear()

        'Message filters are Application-wide, not owned/disposed automatically
        'with the form. Always remove the previous instance before rebuilding the
        'interface, then install exactly one filter for this XtraTabControl.
        RemoveInterfaceScrollRediverter()

        InterfaceScrollRediverter =
            New ScrollRediverter(XtraTabControlNewGIT)

        Application.AddMessageFilter(InterfaceScrollRediverter)

        XtraTabPageCount = -1
        AbovoTabPagesCount = -1

        ReDim XtraTabPages(DataPres.Sections.Length - 1)
        ReDim TPs(DataPres.Sections.Length - 1)
        ReDim AbovoTabPages(DataPres.Sections.Length - 1)

        SuppressLazyTabEvents = True
        XtraTabControlNewGIT.BeginUpdate()

        Try

            For SectionIndex As Integer = 0 To DataPres.Sections.Length - 1

                Dim Section As PresentationSection = DataPres.Sections(SectionIndex)
                Dim CtlName As String = "XtraTabPage" & (SectionIndex + 1).ToString
                Dim TPName As String = "TablePanel" & (SectionIndex + 1).ToString

                Dim CurrTabPage As XtraTabPage = TryCast(XtraTabControlNewGIT.Controls(CtlName), XtraTabPage)

                If CurrTabPage Is Nothing AndAlso SectionIndex < XtraTabControlNewGIT.TabPages.Count Then
                    CurrTabPage = XtraTabControlNewGIT.TabPages(SectionIndex)
                End If

                If CurrTabPage Is Nothing Then Continue For

                XtraTabPages(SectionIndex) = CurrTabPage
                XtraTabPageCount = SectionIndex

                AbovoTabPages(SectionIndex) = New AbovoTabPage With {
                    .Index = SectionIndex,
                    .TabPage = CurrTabPage
                }
                AbovoTabPagesCount = SectionIndex

                CurrTabPage.PageVisible = True
                CurrTabPage.AutoScroll = True
                CurrTabPage.Text = " " & Section.Name & " "

                'The designer contains placeholder TablePanels. They are only layout
                'scaffolding now; real section content is created lazily on selection.
                Dim CurrDummyTP As TablePanel = TryCast(CurrTabPage.Controls(TPName), TablePanel)
                If CurrDummyTP IsNot Nothing Then CurrDummyTP.Visible = False

                Dim State As New InterfaceSectionRuntimeState With {
                    .SectionID = SectionIndex,
                    .IsBuilt = False,
                    .IsDirty = True,
                    .NeedsPresentationRedefinition = False
                }

                SectionRuntimeStates(SectionIndex) = State
                RegisterSectionDependencies(SectionIndex, Section)

            Next

        Finally
            XtraTabControlNewGIT.EndUpdate()
            SuppressLazyTabEvents = False
        End Try

        Dim InitialSection As Integer = XtraTabControlNewGIT.SelectedTabPageIndex
        If InitialSection < 0 Then InitialSection = 0

        EnsureSectionBuilt(InitialSection)

        EndTimer()

    End Sub

    Private Sub RegisterSectionDependencies(ByVal SectionIndex As Integer,
                                            ByVal Section As PresentationSection)

        If Section Is Nothing Then Return

        Dim State As InterfaceSectionRuntimeState = Nothing

        If Not SectionRuntimeStates.TryGetValue(SectionIndex, State) Then
            State = New InterfaceSectionRuntimeState With {.SectionID = SectionIndex}
            SectionRuntimeStates(SectionIndex) = State
        End If

        State.SourceWorksheets.Clear()

        For Each SectionElement In Section.SectionElements

            If SectionElement.ControlSourceIndex < 0 OrElse
               SectionElement.ControlSourceIndex >= DataPres.DataSets.Count Then Continue For

            AddDataSetDependencies(State, DataPres.DataSets(SectionElement.ControlSourceIndex))

        Next

        If ExcelModels(ModelID).InterfaceDependencies IsNot Nothing Then
            ExcelModels(ModelID).InterfaceDependencies.RegisterSection(Me, SectionIndex, State.SourceWorksheets)
        End If

    End Sub

    Private Sub RegisterDataSetDependencies(ByVal SectionIndex As Integer,
                                             ByVal DataSet As DataCellRange)

        If DataSet Is Nothing Then Return

        Dim State As InterfaceSectionRuntimeState = Nothing

        If Not SectionRuntimeStates.TryGetValue(SectionIndex, State) Then
            State = New InterfaceSectionRuntimeState With {.SectionID = SectionIndex}
            SectionRuntimeStates(SectionIndex) = State
        End If

        AddDataSetDependencies(State, DataSet)

        If ExcelModels(ModelID).InterfaceDependencies IsNot Nothing Then
            ExcelModels(ModelID).InterfaceDependencies.RegisterSection(Me, SectionIndex, State.SourceWorksheets)
        End If

    End Sub

    Private Sub AddDataSetDependencies(ByVal State As InterfaceSectionRuntimeState,
                                       ByVal DataSet As DataCellRange)

        If State Is Nothing OrElse DataSet Is Nothing Then Return

        If Not String.IsNullOrWhiteSpace(DataSet.SourceWorksheet) Then
            State.SourceWorksheets.Add(DataSet.SourceWorksheet)
        End If

        If DataSet.DataRows IsNot Nothing Then
            For Each DataRow In DataSet.DataRows
                If DataRow.DataCells Is Nothing Then Continue For
                For Each DP As CellDataPoint In DataRow.DataCells
                    If DP IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(DP.SourceSheet) Then
                        State.SourceWorksheets.Add(DP.SourceSheet)
                    End If
                Next
            Next
        End If

        'Structural named ranges can determine the shape of a dataset even when
        'the displayed cells themselves are sourced from another worksheet.
        AddNamedRangeDependency(State, DataSet.RowExpandByNR)
        AddNamedRangeDependency(State, DataSet.ColExpandByNR)
        AddNamedRangeDependency(State, DataSet.RepeatingNR)
        AddNamedRangeDependency(State, DataSet.DefaultDataNR)
        AddNamedRangeDependency(State, DataSet.LiveGridSourceName)

        If DataSet.DataColumns IsNot Nothing Then
            For Each DataColumn In DataSet.DataColumns
                If DataColumn.ColumnTag Is Nothing Then Continue For
                AddNamedRangeDependency(State, DataColumn.ColumnTag.ActionNR)
                AddNamedRangeDependency(State, DataColumn.ColumnTag.RepeatingNR)
            Next
        End If

    End Sub

    Private Sub AddNamedRangeDependency(ByVal State As InterfaceSectionRuntimeState,
                                        ByVal NamedRange As String)

        If State Is Nothing OrElse String.IsNullOrWhiteSpace(NamedRange) Then Return
        If ExcelModels(ModelID).InterfaceDependencies Is Nothing Then Return

        Dim WorksheetName As String =
            ExcelModels(ModelID).InterfaceDependencies.GetNamedRangeWorksheetName(NamedRange)

        If Not String.IsNullOrWhiteSpace(WorksheetName) Then State.SourceWorksheets.Add(WorksheetName)

    End Sub

    Private Sub EnsureSectionBuilt(ByVal SectionIndex As Integer,
                                   Optional ByVal DeferFinalViewportLayout As Boolean = False)

        If SectionIndex < 0 OrElse SectionIndex >= DataPres.Sections.Length Then Return

        Dim State As InterfaceSectionRuntimeState = Nothing

        If Not SectionRuntimeStates.TryGetValue(SectionIndex, State) Then
            State = New InterfaceSectionRuntimeState With {.SectionID = SectionIndex}
            SectionRuntimeStates(SectionIndex) = State
        End If

        If State.IsBuilt AndAlso
           Not State.IsDirty AndAlso
           Not State.NeedsPresentationRedefinition Then Return

        BuildSection(
            SectionIndex,
            State.NeedsPresentationRedefinition,
            DeferFinalViewportLayout)

    End Sub

    Private Sub BuildSection(ByVal SectionIndex As Integer,
                             Optional ByVal RedefinePresentation As Boolean = False,
                             Optional ByVal DeferFinalViewportLayout As Boolean = False)

        If SectionIndex < 0 OrElse SectionIndex >= DataPres.Sections.Length Then Return

        Dim XTP As XtraTabPage = XtraTabControlNewGIT.TabPages(SectionIndex)
        If XTP Is Nothing Then Return

        'Build the whole section while XtraTabControl painting is locked.  This is
        'particularly useful for lazy tabs: the user sees the completed page rather
        'than the intermediate empty/disposed state.  BeginUpdate/EndUpdate are
        'counter based, so this is also safe when the caller already holds an update.
        XtraTabControlNewGIT.BeginUpdate()
        XTP.SuspendLayout()

        Try

            If RedefinePresentation Then DataPres.RedefineInterfaceSection(SectionIndex)

            Dim ThisSection As PresentationSection = DataPres.Sections(SectionIndex)

            UnloadSection(SectionIndex)

            AbovoTabPages(SectionIndex) = New AbovoTabPage With {
                .Index = SectionIndex,
                .TabPage = XTP
            }
            CurrentAbovoTabPage = AbovoTabPages(SectionIndex)

            'Scroll ownership is intentionally NOT shared with the page.
            'XtraTabPage owns interface scrolling; GridControl/VGridControl
            'own scrolling of their data. The TablePanel is layout only.
            Dim NewTP As New TablePanel With {
                .Padding = DefaultTablePanelPadding,
                .AutoScroll = False,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Dock = DockStyle.Top,
                .Visible = False
            }

            NewTP.SuspendLayout()

            NewTP.Columns.Clear()
            NewTP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 2))
            NewTP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 1))
            NewTP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 5))
            NewTP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Absolute, 100))

            TPs(SectionIndex) = NewTP

            'Create the child controls while the TablePanel is hidden.
            GetSectionControlCollection(ThisSection, Me, SectionIndex, NewTP)
            XTP.Controls.Add(NewTP)

            'IMPORTANT: let TablePanel resolve its rows/columns and child bounds
            'BEFORE asking DevExpress controls for best sizes.  The previous order
            'called ApplySectionFontAndGridLayout while NewTP was still suspended,
            'so Calc/ViewInfo values reflected pre-layout dimensions.
            NewTP.ResumeLayout(True)
            NewTP.PerformLayout()

            ApplySectionFontAndGridLayout(NewTP)
            NewTP.PerformLayout()

            NewTP.Visible = True
            NewTP.BringToFront()

            Dim State As InterfaceSectionRuntimeState = Nothing
            If Not SectionRuntimeStates.TryGetValue(SectionIndex, State) Then
                State = New InterfaceSectionRuntimeState With {.SectionID = SectionIndex}
                SectionRuntimeStates(SectionIndex) = State
            End If

            State.IsBuilt = True
            State.IsDirty = False
            State.NeedsPresentationRedefinition = False

            RegisterSectionDependencies(SectionIndex, ThisSection)

        Finally

            XTP.ResumeLayout(True)
            XtraTabControlNewGIT.EndUpdate()

            'The first sizing pass occurs while XTP is suspended and the
            'XtraTabControl is still inside BeginUpdate. On the large display
            'that pass sees an old page geometry (~1236 px high).
            '
            'For a normal/initial build, run the authoritative pass now.
            '
            'For a lazy tab build triggered by SelectedPageChanging, DO NOT force
            'the whole XtraTabControl through another layout while the OLD tab is
            'still selected. That was the source of the visible "jump down":
            'the current page could repaint/reflow before the requested tab had
            'actually become selected.
            '
            'SelectedPageChanged already performs the authoritative sizing pass
            'after the new page has its real viewport, so defer to that event.
            If Not DeferFinalViewportLayout Then

                XtraTabControlNewGIT.PerformLayout()
                XTP.PerformLayout()

                Dim FinalTP As TablePanel = Nothing

                If TPs IsNot Nothing AndAlso
                   SectionIndex >= 0 AndAlso
                   SectionIndex < TPs.Length Then

                    FinalTP = TPs(SectionIndex)

                End If

                If FinalTP IsNot Nothing AndAlso
                   Not FinalTP.IsDisposed Then

                    FinalTP.PerformLayout()
                    ApplySectionFontAndGridLayout(FinalTP)
                    FinalTP.PerformLayout()

                End If

            End If

#If DEBUG Then
#End If

        End Try

    End Sub

    Private Sub DumpSectionGridLayoutDiagnostics(
        ByVal SectionIndex As Integer)

#If DEBUG Then
        Try

            If SectionIndex < 0 OrElse
               SectionIndex >= XtraTabControlNewGIT.TabPages.Count Then Return

            Dim XTP As XtraTabPage =
                XtraTabControlNewGIT.TabPages(SectionIndex)

            Dim TP As TablePanel = Nothing

            If TPs IsNot Nothing AndAlso
               SectionIndex >= 0 AndAlso
               SectionIndex < TPs.Length Then

                TP = TPs(SectionIndex)

            End If


            If ParentGroupForm IsNot Nothing Then
            End If


            If XTP IsNot Nothing Then
            End If

            If TP IsNot Nothing Then

                For RowIndex As Integer = 0 To TP.Rows.Count - 1
                Next
            End If

            If TP IsNot Nothing Then

                For Each VG As VGridControl In
                    FindChildControls(Of VGridControl)(TP)

                    DumpOneVGridLayoutDiagnostic(
                        VG,
                        XTP,
                        TP)

                Next

                For Each GC As GridControl In
                    FindChildControls(Of GridControl)(TP)

                    DumpOneGridLayoutDiagnostic(
                        GC,
                        XTP,
                        TP)

                Next

            End If


        Catch ex As Exception


        End Try
#End If

    End Sub

    Private Sub DumpOneVGridLayoutDiagnostic(
        ByVal VG As VGridControl,
        ByVal XTP As XtraTabPage,
        ByVal TP As TablePanel)

#If DEBUG Then
        If VG Is Nothing Then Return

        Dim GridTopInTab As Integer = -1
        Dim GridBottomInTab As Integer = -1
        Dim RemainingTabHeight As Integer = -1

        Try
            Dim P As Point =
                XTP.PointToClient(
                    VG.PointToScreen(Point.Empty))

            GridTopInTab = P.Y
            GridBottomInTab = P.Y + VG.Height
            RemainingTabHeight =
                XTP.ClientSize.Height - GridTopInTab
        Catch
        End Try

        Dim LayoutTag As VGridLayoutTag =
            TryCast(VG.Tag, VGridLayoutTag)

        Dim TableRowIndex As Integer =
            If(LayoutTag Is Nothing,
               -1,
               LayoutTag.TableRowIndex)

        Dim TableRowHeight As Single = -1
        Dim TableRowStyle As String = "?"

        If TableRowIndex >= 0 AndAlso
           TableRowIndex < TP.Rows.Count Then

            TableRowHeight =
                TP.Rows(TableRowIndex).Height

            TableRowStyle =
                TP.Rows(TableRowIndex).Style.ToString

        End If


        DumpControlParentChain(VG)

#End If

    End Sub

    Private Sub DumpOneGridLayoutDiagnostic(
        ByVal GC As GridControl,
        ByVal XTP As XtraTabPage,
        ByVal TP As TablePanel)

#If DEBUG Then
        If GC Is Nothing Then Return

        Dim GridTopInTab As Integer = -1
        Dim RemainingTabHeight As Integer = -1

        Try
            Dim P As Point =
                XTP.PointToClient(
                    GC.PointToScreen(Point.Empty))

            GridTopInTab = P.Y
            RemainingTabHeight =
                XTP.ClientSize.Height - GridTopInTab
        Catch
        End Try


#End If

    End Sub

    Private Sub DumpControlParentChain(
        ByVal StartControl As Control)

#If DEBUG Then
        Dim CurrentControl As Control =
            StartControl

        Dim Level As Integer = 0

        While CurrentControl IsNot Nothing AndAlso Level < 12


            CurrentControl =
                CurrentControl.Parent

            Level += 1

        End While
#End If

    End Sub

    Private Sub ApplySectionFontAndGridLayout(ByVal RootControl As Control)

        If RootControl Is Nothing OrElse RootControl.IsDisposed Then Return

        Dim CurrentScaleFactor As Single

        If ParentGroupForm Is Nothing Then
            CurrentScaleFactor = Me.Width / 1700
        Else
            CurrentScaleFactor = ParentGroupForm.Width / 1700
        End If

        Dim NewFont As Font = GetFont("Small", CurrentScaleFactor)

        ApplyFontToControlTree(RootControl, NewFont)

        For Each GC As GridControl In FindChildControls(Of GridControl)(RootControl)

            If GC Is Nothing OrElse GC.IsDisposed Then Continue For

            Dim GV As GridView = TryCast(GC.MainView, GridView)
            If GV Is Nothing Then Continue For

            GV.CheckLoaded()
            GV.BeginUpdate()

            Try

                For Each ap As AppearanceObject In GV.Appearance
                    ap.Font = NewFont
                Next

                Dim ViewTag As GridViewTag = TryCast(GV.Tag, GridViewTag)

                If ViewTag IsNot Nothing AndAlso Not ViewTag.HaveProcessedColumns Then
                    Formatter.ProcessGVColumWidths(GV, Me)
                End If

                'Use the full responsive fit at build/layout time. This sizes
                'columns to their actual header/data content and then allows the
                'GridControl itself to grow into useful available workspace.
                ApplyResponsiveGridWidth(
                    GC,
                    GV,
                    RootControl,
                    False)

                Dim GVI As GridViewInfo = TryCast(GV.GetViewInfo(), GridViewInfo)

                If GVI IsNot Nothing Then
                    IdealGridRowHeight = GVI.MinRowHeight
                End If

            Finally
                GV.EndUpdate()
            End Try

        Next

        'Size every GridControl from DevExpress' own calculated content size.
        'This is intentionally view-agnostic: it covers GridView AND BandedGridView
        'and includes headers, rows, footer panels and scrollbars.
        For Each GC As GridControl In FindChildControls(Of GridControl)(RootControl)

            If GC Is Nothing OrElse GC.IsDisposed Then Continue For

            GC.ForceInitialize()
            GC.PerformLayout()

            Dim MaxBestWidth As Integer = Math.Max(1, GC.ClientSize.Width)

            If MaxBestWidth <= 1 Then
                MaxBestWidth = Math.Max(1, RootControl.ClientSize.Width)
            End If

            Dim AvailableGridHeight As Integer =
                GetAvailableGridViewportHeight(
                    GC,
                    RootControl)

            Dim MaxGridSize As New Size(
                MaxBestWidth,
                AvailableGridHeight)

            'DevExpress recommends calculating once, updating layout, then
            'calculating again with scrollbar information included.
            Dim BestGridSize As Size =
                GC.CalcBestSize(
                    MaxGridSize,
                    False)

            If BestGridSize.Height > 0 Then
                GC.Height = BestGridSize.Height
            End If

            If GC.MainView IsNot Nothing Then
                GC.MainView.LayoutChanged()
            End If

            GC.PerformLayout()

            BestGridSize =
                GC.CalcBestSize(
                    MaxGridSize,
                    True)

            If BestGridSize.Height > 0 Then
                GC.Height = BestGridSize.Height
            End If

#If DEBUG Then
#End If

        Next

        'VGrid BestFit must happen AFTER the final font is applied and AFTER the
        'TablePanel has resolved its actual available width.
        For Each VG As VGridControl In FindChildControls(Of VGridControl)(RootControl)

            If VG Is Nothing OrElse VG.IsDisposed Then Continue For

            VG.Font = NewFont
            VG.ForceInitialize()
            VG.BestFit()

            'BestFit is content-driven and can become much too aggressive on
            'wide VGrids such as Funding, particularly on 4K/high-DPI displays:
            'DevExpress reduces the common RecordWidth until every facility is
            'visible, which leaves headers and values heavily ellipsised.
            '
            'Treat BestFit as the preferred size, then apply DPI/font-aware
            'MINIMUMS.  We never reduce a width selected by BestFit (or an
            'explicitly wider specialised VGrid such as Joint Ventures).
            '
            'If the resulting record set is wider than the viewport, the existing
            'horizontal-scroll policy below will handle it.
            Dim MinimumRecordWidth As Integer =
                GetMinimumVGridRecordWidth(
                    VG,
                    NewFont)

            If VG.RecordWidth < MinimumRecordWidth Then
                VG.RecordWidth = MinimumRecordWidth
            End If

            Dim MinimumRowHeaderWidth As Integer =
                GetMinimumVGridRowHeaderWidth(
                    VG,
                    NewFont)

            If VG.RowHeaderWidth < MinimumRowHeaderWidth Then
                VG.RowHeaderWidth = MinimumRowHeaderWidth
            End If

            'For VGrid interfaces the XtraTabPage owns VERTICAL scrolling.
            'Keep horizontal scrolling available for wide record sets, but never
            'show a competing VGrid vertical scrollbar.
            VG.ScrollVisibility =
                DevExpress.XtraVerticalGrid.ScrollVisibility.Horizontal

            Dim AvailableVGridWidth As Integer =
                Math.Max(1, RootControl.ClientSize.Width -
                            RootControl.Padding.Horizontal)

            Dim PreferredVGridWidth As Integer =
                GetPreferredVGridContentWidth(
                    VG,
                    AvailableVGridWidth)

            VG.Width = PreferredVGridWidth

            'VGrid is inverted: DataCellRange.RowCount is the number of RECORDS
            '(displayed across the control), not the number of visible VGrid rows.
            'The old FillSize * average-row-height calculation therefore used the
            'wrong dimension for Funding and similar Merge/Pivot VGrids.
            '
            'After BestFit, sum the ACTUAL visible VGrid row heights (including
            'expanded categories/children), then cap that content height to the
            'remaining visible tab-page viewport.  If there is more data than
            'fits, the VGrid fills the available screen height and scrolls
            'internally; if there is less, it remains content-sized.
            Dim AvailableVGridHeight As Integer =
                GetAvailableGridViewportHeight(
                    VG,
                    RootControl)

            Dim PreferredVGridHeight As Integer =
                GetPreferredVGridContentHeight(VG)

            'The row-height sum is based on public BaseRow metrics. DevExpress
            'also consumes a small amount of internal chrome/rounding that is not
            'fully represented by those row heights. Since the VGrid vertical
            'scrollbar is intentionally hidden, reserve one extra row-height-sized
            'buffer so the final synthetic footer row cannot be clipped.
            Dim VGridVerticalSafetyReserve As Integer =
                Math.Max(
                    24,
                    IdealGridRowHeight + (2 * DefaultGridCellPadding))

            PreferredVGridHeight += VGridVerticalSafetyReserve

            'VGrid scrolling policy:
            '
            'The XtraTabPage is the SINGLE vertical scroll owner for VGrid-based
            'interfaces.  Therefore the VGrid must be tall enough to contain ALL
            'of its visible rows; otherwise DevExpress creates an inner vertical
            'scroll range and lower category/footer rows (such as Commitment Fees
            'and its Add rows action) become a second, nested scroll surface.
            '
            'At the same time, when the VGrid content is shorter than the visible
            'page, retain the requested full-workspace appearance.
            '
            'So:
            '   short VGrid -> fill remaining viewport
            '   tall VGrid  -> expand to full content height; XtraTabPage scrolls
            Dim AppliedVGridHeight As Integer =
                Math.Max(
                    AvailableVGridHeight,
                    PreferredVGridHeight)

            VG.Height = AppliedVGridHeight

            'The parent TablePanel row is the controlling geometry. With an
            'AutoSize row DevExpress ignores Row.Height and derives the row from
            'content, so changing only VG.Height cannot make the cell fill the
            'viewport. Convert this VGrid's owning row to Absolute and set the
            'same viewport height explicitly.
            Dim ThisVGridLayout As VGridLayoutTag =
                TryCast(VG.Tag, VGridLayoutTag)

            Dim VGridTablePanel As TablePanel =
                TryCast(VG.Parent, TablePanel)

            If ThisVGridLayout IsNot Nothing AndAlso
               VGridTablePanel IsNot Nothing AndAlso
               ThisVGridLayout.TableRowIndex >= 0 AndAlso
               ThisVGridLayout.TableRowIndex < VGridTablePanel.Rows.Count Then

                VGridTablePanel.Rows(ThisVGridLayout.TableRowIndex).Style =
                    TablePanelEntityStyle.Absolute

                VGridTablePanel.Rows(ThisVGridLayout.TableRowIndex).Height =
                    AppliedVGridHeight

            End If

            VG.MinimumSize =
                New Size(
                    VG.MinimumSize.Width,
                    AppliedVGridHeight)

            'A VGrid command host uses Tag to identify the VGrid it belongs to.
            '
            'Do NOT rely on the host Panel's own width for alignment: DevExpress
            'TablePanel can stretch a child to the full spanned cell width on very
            'wide displays.  Instead position the command buttons explicitly
            'against the ACTUAL VGrid content width.
            For Each CommandHost As System.Windows.Forms.Panel In
                FindChildControls(Of System.Windows.Forms.Panel)(RootControl)

                If CommandHost Is Nothing OrElse CommandHost.IsDisposed Then Continue For
                If Not Object.ReferenceEquals(CommandHost.Tag, VG) Then Continue For

                Dim RightEdge As Integer =
                    Math.Max(
                        CommandHost.Padding.Left,
                        PreferredVGridWidth - CommandHost.Padding.Right)

                'Commands are stored in XML order. Position from right to left so
                'one button sits at the VGrid's top-right and multiple exceptional
                'VGrid commands remain supported.
                For ButtonIndex As Integer =
                    CommandHost.Controls.Count - 1 To 0 Step -1

                    Dim CommandButton As Control =
                        CommandHost.Controls(ButtonIndex)

                    If CommandButton Is Nothing OrElse
                       CommandButton.IsDisposed Then Continue For

                    Dim ButtonLeft As Integer =
                        Math.Max(
                            CommandHost.Padding.Left,
                            RightEdge - CommandButton.Width)

                    CommandButton.Location =
                        New Point(
                            ButtonLeft,
                            CommandHost.Padding.Top)

                    RightEdge =
                        ButtonLeft -
                        Math.Max(4, CommandButton.Margin.Left + CommandButton.Margin.Right)

                Next

#If DEBUG Then
#End If

            Next

#If DEBUG Then
#End If

        Next

        'A DockStyle.Top, AutoSize TablePanel is the page's content surface.
        'Grid/VGrid sizing above can change absolute row heights, so force the
        'panel to recalculate its own content height before XtraTabPage computes
        'its AutoScroll extent.
        Dim LayoutTP As TablePanel =
            TryCast(RootControl, TablePanel)

        If LayoutTP IsNot Nothing Then
            LayoutTP.PerformLayout()
        End If

        If RootControl IsNot Nothing AndAlso
           RootControl.Parent IsNot Nothing Then

            RootControl.Parent.PerformLayout()

        End If

    End Sub

    Private Function GetGridColumnConfiguredMinimumWidth(
        ByVal GV As GridView,
        ByVal Col As GridColumn) As Integer

        If GV Is Nothing OrElse Col Is Nothing Then Return 0

        Dim ColTag As DataColumnTag =
            TryCast(Col.Tag, DataColumnTag)

        If ColTag Is Nothing OrElse
           ColTag.MinimumWidthChars <= 0 Then

            Return 0

        End If

        Dim MeasureFont As Font = GV.Appearance.Row.Font

        If MeasureFont Is Nothing Then
            MeasureFont = Me.Font
        End If

        'Use a deliberately wide representative character rather than "0".
        'This gives a stable minimum for empty text fields without turning a
        'character count into a fragile fixed-pixel definition.
        Dim SampleText As String =
            New String("M"c, ColTag.MinimumWidthChars)

        Dim MeasuredSize As Size =
            TextRenderer.MeasureText(
                SampleText,
                MeasureFont,
                New Size(Integer.MaxValue, Integer.MaxValue),
                TextFormatFlags.NoPadding Or TextFormatFlags.SingleLine)

        Return MeasuredSize.Width + (2 * DefaultGridCellPadding) + 8

    End Function

    Private Sub ApplyConfiguredGridColumnMinimumWidths(ByVal GV As GridView)

        If GV Is Nothing Then Exit Sub

        For Each Col As GridColumn In GV.Columns

            If Not Col.Visible Then Continue For

            Dim ConfiguredMinimum As Integer =
                GetGridColumnConfiguredMinimumWidth(
                    GV,
                    Col)

            If ConfiguredMinimum > 0 Then

                Col.MinWidth =
                    Math.Max(
                        Col.MinWidth,
                        ConfiguredMinimum)

                If Col.Width < ConfiguredMinimum Then
                    Col.Width = ConfiguredMinimum
                End If

            End If

        Next

    End Sub

    Private Sub ApplyRepeatingFamilyColumnWidths(ByVal GV As GridView)

        If GV Is Nothing Then Exit Sub

        'Columns generated from the same DataFieldDefinition RepeatsByNR family
        'carry the same DataColumnTag.RepeatingNR.  BestFit each column first,
        'then normalise every visible member of that family to the widest member.
        '
        'This gives repeated columns a stable visual rhythm (for example
        'Rent Assumptions > Voids) without forcing unrelated columns to share
        'the same width.
        Dim FamilyWidths As New Dictionary(Of String, Integer)(
            StringComparer.OrdinalIgnoreCase)

        Dim FamilyCounts As New Dictionary(Of String, Integer)(
            StringComparer.OrdinalIgnoreCase)

        For Each Col As GridColumn In GV.Columns

            If Not Col.Visible Then Continue For

            Dim ColTag As DataColumnTag =
                TryCast(Col.Tag, DataColumnTag)

            If ColTag Is Nothing OrElse
               String.IsNullOrWhiteSpace(ColTag.RepeatingNR) Then

                Continue For

            End If

            Dim FamilyKey As String =
                ColTag.RepeatingNR.Trim()

            If Not FamilyWidths.ContainsKey(FamilyKey) Then
                FamilyWidths.Add(FamilyKey, Col.Width)
                FamilyCounts.Add(FamilyKey, 1)
            Else
                FamilyWidths(FamilyKey) =
                    Math.Max(
                        FamilyWidths(FamilyKey),
                        Col.Width)

                FamilyCounts(FamilyKey) += 1
            End If

        Next

        For Each Col As GridColumn In GV.Columns

            If Not Col.Visible Then Continue For

            Dim ColTag As DataColumnTag =
                TryCast(Col.Tag, DataColumnTag)

            If ColTag Is Nothing OrElse
               String.IsNullOrWhiteSpace(ColTag.RepeatingNR) Then

                Continue For

            End If

            Dim FamilyKey As String =
                ColTag.RepeatingNR.Trim()

            'A single column is not a repeating visual family yet, so leave its
            'individually fitted width alone.
            If Not FamilyCounts.ContainsKey(FamilyKey) OrElse
               FamilyCounts(FamilyKey) < 2 Then

                Continue For

            End If

            Dim FamilyWidth As Integer =
                FamilyWidths(FamilyKey)

            Col.MinWidth =
                Math.Max(
                    Col.MinWidth,
                    FamilyWidth)

            Col.Width = FamilyWidth

        Next

#If DEBUG Then

        For Each FamilyKey As String In FamilyWidths.Keys

            If FamilyCounts(FamilyKey) > 1 Then


            End If

        Next

#End If

    End Sub

    Private Sub AutoFitGridAfterDataRefresh(ByVal GC As GridControl)

        If GC Is Nothing OrElse GC.IsDisposed Then Exit Sub

        Dim GV As GridView =
            TryCast(GC.MainView, GridView)

        If GV Is Nothing Then Exit Sub

        Dim RootControl As Control = GC.Parent

        While RootControl IsNot Nothing AndAlso
              Not TypeOf RootControl Is TablePanel AndAlso
              RootControl.Parent IsNot Nothing

            RootControl = RootControl.Parent

        End While

        If RootControl Is Nothing Then
            RootControl = GC.Parent
        End If

        GV.CheckLoaded()
        GV.BeginUpdate()

        Try

            'Refresh-time fitting is deliberately GROW-ONLY. A user may have
            'manually widened a column and we should not collapse it again simply
            'because a later refresh contains shorter text.
            ApplyResponsiveGridWidth(
                GC,
                GV,
                RootControl,
                True)

        Finally
            GV.EndUpdate()
        End Try

        GC.PerformLayout()

        If RootControl IsNot Nothing Then
            RootControl.PerformLayout()
        End If

    End Sub

    Private Sub ApplyResponsiveGridWidth(ByVal GC As GridControl,
                                         ByVal GV As GridView,
                                         ByVal RootControl As Control,
                                         ByVal GrowOnly As Boolean)

        If GC Is Nothing OrElse GV Is Nothing Then Exit Sub
        If GC.IsDisposed Then Exit Sub

        Dim PreviousWidths As New Dictionary(Of GridColumn, Integer)
        Dim OriginalGridWidth As Integer = GC.Width

        If GrowOnly Then

            For Each Col As GridColumn In GV.Columns

                If Col.Visible Then
                    PreviousWidths(Col) = Col.Width
                End If

            Next

        End If

        'Measure the grid at its NATURAL content width.
        '
        'ColumnAutoWidth must be OFF before BestFitColumns. If it is left on,
        'DevExpress can redistribute spare client width across the columns; on a
        '5K display that can make a modest grid fill almost the whole screen.
        GV.OptionsView.ColumnAutoWidth = False
        GV.OptionsView.BestFitMaxRowCount = -1
        GV.BestFitColumns()

        'BestFit is data-driven. On an empty/sparse grid there may be nothing
        'wide enough to establish a useful editing surface, so honour any
        'Structure.xml MinWidthChars after measuring actual content.
        ApplyConfiguredGridColumnMinimumWidths(GV)

        'RepeatsByNR-generated columns are one visual family. Once each member
        'has been fitted to its own current data/configured minimum, make the
        'family consistently as wide as its widest member.
        ApplyRepeatingFamilyColumnWidths(GV)

        If GrowOnly Then

            For Each Pair As KeyValuePair(Of GridColumn, Integer) In PreviousWidths

                If Pair.Key IsNot Nothing AndAlso
                   Pair.Key.Visible AndAlso
                   Pair.Key.Width < Pair.Value Then

                    Pair.Key.Width = Pair.Value

                End If

            Next

        End If

        'GrowOnly can restore a previously/manual widened family member above
        'the freshly fitted family width. Re-normalise once more so all members
        'remain equal, using that restored wider member as the new family width.
        ApplyRepeatingFamilyColumnWidths(GV)

        Dim NaturalContentWidth As Integer = 0

        For Each Col As GridColumn In GV.Columns

            If Col.Visible Then
                NaturalContentWidth += Math.Max(1, Col.Width)
            End If

        Next

        'Only allow for genuine grid chrome. There is intentionally no
        'screen-scaled minimum width: unused 5K workspace should stay empty.
        NaturalContentWidth +=
            Math.Max(36, GV.IndicatorWidth) +
            SystemInformation.VerticalScrollBarWidth +
            12

        Dim AvailableWidth As Integer =
            GetAvailableGridViewportWidth(
                GC,
                RootControl)

        Dim DesiredWidth As Integer =
            Math.Min(
                AvailableWidth,
                Math.Max(1, NaturalContentWidth))

        If GrowOnly Then

            'Refresh/paste can grow the grid for longer data, but should not
            'collapse a control that the user has manually widened.
            DesiredWidth =
                Math.Min(
                    AvailableWidth,
                    Math.Max(
                        GC.Width,
                        DesiredWidth))

        End If

        If DesiredWidth > 0 Then
            GC.Width = DesiredWidth
        End If

        'Keep natural fitted widths. If content genuinely exceeds the
        'available page width, the GridControl will use horizontal scrolling
        'rather than compressing the columns back down.
        GV.OptionsView.ColumnAutoWidth = False

        'ColumnInplaceEditorHelper / ColumnButtonExtender calculate their hit
        'rectangles from GridViewInfo.ColumnsInfo. Fix 47 can change several
        'RepeatsByNR column widths AFTER those helpers have subscribed.
        '
        'On complex/lazy-built banded grids DevExpress can continue painting with
        'the new widths while ColumnsInfo still contains the previous geometry.
        'The header therefore LOOKS correct but a click is tested against the old
        'rectangle and appears to do nothing.
        '
        'Force one authoritative view-info rebuild after all width/minimum/family
        'changes are complete, then invalidate the headers. This keeps painting
        'and mouse hit-testing on exactly the same geometry.
        GV.LayoutChanged()
        GV.Invalidate()

        If GC.Width <> OriginalGridWidth Then
            GC.PerformLayout()
        End If

#If DEBUG Then
#End If

    End Sub

    Private Function GetAvailableGridViewportWidth(
        ByVal GridControlToSize As Control,
        ByVal RootControl As Control) As Integer

        If GridControlToSize Is Nothing Then Return 600

        Dim Viewport As Control = Nothing
        Dim ParentControl As Control = GridControlToSize.Parent

        While ParentControl IsNot Nothing

            If TypeOf ParentControl Is XtraTabPage Then
                Viewport = ParentControl
                Exit While
            End If

            ParentControl = ParentControl.Parent

        End While

        If Viewport Is Nothing Then
            Viewport = RootControl
        End If

        If Viewport Is Nothing Then
            Return Math.Max(600, GridControlToSize.Width)
        End If

        Dim GridLeftInViewport As Integer = 0

        Try
            GridLeftInViewport =
                Viewport.PointToClient(
                    GridControlToSize.PointToScreen(
                        Point.Empty)).X
        Catch
            GridLeftInViewport = GridControlToSize.Left
        End Try

        Dim RightReserve As Integer =
            Math.Max(
                12,
                Viewport.Padding.Right +
                SystemInformation.VerticalScrollBarWidth)

        Dim RetWidth As Integer =
            Viewport.ClientSize.Width -
            Math.Max(0, GridLeftInViewport) -
            RightReserve

        Return Math.Max(1, RetWidth)

    End Function

    Private Function GetAvailableGridViewportHeight(
        ByVal GridControlToSize As Control,
        ByVal RootControl As Control) As Integer

        If GridControlToSize Is Nothing Then Return 300

        Dim Viewport As Control = Nothing
        Dim SearchControl As Control = RootControl

        'The XtraTabPage is the actual visible section viewport.  Using the
        'monitor/screen height directly would be wrong when the application is
        'not maximised or when toolbars/tab headers consume part of the window.
        While SearchControl IsNot Nothing

            If TypeOf SearchControl Is XtraTabPage Then
                Viewport = SearchControl
                Exit While
            End If

            SearchControl = SearchControl.Parent

        End While

        If Viewport Is Nothing AndAlso RootControl IsNot Nothing Then
            Viewport = RootControl.Parent
        End If

        If Viewport Is Nothing Then
            Viewport = Me
        End If

        Dim GridTopInViewport As Integer

        Try

            Dim GridScreenPoint As Point =
                GridControlToSize.PointToScreen(Point.Empty)

            GridTopInViewport =
                Viewport.PointToClient(GridScreenPoint).Y

        Catch

            GridTopInViewport = GridControlToSize.Top

        End Try

        Dim BottomReserve As Integer =
            Math.Max(
                12,
                DefaultTablePanelPadding.Bottom + 8)

        Dim AvailableHeight As Integer =
            Viewport.ClientSize.Height -
            GridTopInViewport -
            BottomReserve

        'Never return more than the viewport itself.  This also protects against
        'negative translated Y values when a scrollable page is already scrolled.
        AvailableHeight =
            Math.Min(
                AvailableHeight,
                Viewport.ClientSize.Height - BottomReserve)

        'Keep enough room for a useful control even in a very small window.
        Return Math.Max(120, AvailableHeight)

    End Function

    Private Function GetPreferredVGridContentHeight(
        ByVal VG As VGridControl) As Integer

        If VG Is Nothing Then Return 120

        Dim TotalHeight As Integer = 0

        For Each TopRow As BaseRow In VG.Rows

            TotalHeight +=
                GetVisibleVGridRowHeight(TopRow)

        Next

        'Small allowance for grid borders/chrome.  If the preferred width is
        'capped by the viewport, DevExpress may also need its horizontal
        'scrollbar; reserve that height rather than clipping the final row.
        'Allow for VGrid borders/focus rectangles and rounding introduced by
        'DevExpress row metrics. A small over-allocation is preferable to an
        'unwanted one-row internal vertical scrollbar.
        TotalHeight += 16

        If VG.Width <
           GetUncappedVGridContentWidth(VG) Then

            TotalHeight +=
                SystemInformation.HorizontalScrollBarHeight

        End If

        Return Math.Max(120, TotalHeight)

    End Function

    Private Function GetVisibleVGridRowHeight(
        ByVal Row As BaseRow) As Integer

        If Row Is Nothing OrElse Not Row.Visible Then Return 0

        Dim TotalHeight As Integer =
            Math.Max(
                BaseRow.MinHeight,
                Row.Height)

        If Row.HasChildren AndAlso Row.Expanded Then

            For Each ChildRow As BaseRow In Row.ChildRows

                TotalHeight +=
                    GetVisibleVGridRowHeight(ChildRow)

            Next

        End If

        Return TotalHeight

    End Function

    Private Function GetUncappedVGridContentWidth(
        ByVal VG As VGridControl) As Integer

        If VG Is Nothing Then Return 0

        Dim PreferredWidth As Long =
            Math.Max(0, VG.RowHeaderWidth)

        If VG.RecordCount > 0 Then

            PreferredWidth +=
                CLng(Math.Max(0, VG.RecordWidth)) *
                CLng(VG.RecordCount)

            If VG.RecordCount > 1 Then

                PreferredWidth +=
                    CLng(Math.Max(0, VG.RecordsInterval)) *
                    CLng(VG.RecordCount - 1)

            End If

        End If

        PreferredWidth += 8

        If PreferredWidth > Integer.MaxValue Then
            Return Integer.MaxValue
        End If

        Return CInt(PreferredWidth)

    End Function

    Private Function GetVGridInterfaceScale(ByVal VG As VGridControl,
                                             ByVal AppliedFont As Font) As Single

        Dim DpiScale As Single = 1.0F
        Dim FontScale As Single = 1.0F

        Try
            Using G As Graphics = VG.CreateGraphics()
                If G IsNot Nothing AndAlso G.DpiX > 0 Then
                    DpiScale = G.DpiX / 96.0F
                End If
            End Using
        Catch
            DpiScale = 1.0F
        End Try

        If AppliedFont IsNot Nothing AndAlso
           SystemFonts.MessageBoxFont IsNot Nothing AndAlso
           SystemFonts.MessageBoxFont.SizeInPoints > 0 Then

            FontScale =
                AppliedFont.SizeInPoints /
                SystemFonts.MessageBoxFont.SizeInPoints

        End If

        'The interface font is already responsive to the host form size, while
        'Windows/DevExpress may additionally DPI-scale the control.  Taking the
        'larger factor gives us a useful physical minimum without multiplying the
        'two factors together and over-inflating widths on a 4K display.
        Return Math.Max(
            1.0F,
            Math.Max(DpiScale, FontScale))

    End Function

    Private Function GetMinimumVGridRecordWidth(ByVal VG As VGridControl,
                                                ByVal AppliedFont As Font) As Integer

        If VG Is Nothing Then Return 120

        Dim Scale As Single =
            GetVGridInterfaceScale(
                VG,
                AppliedFont)

        '120 px is the practical 100%-scale floor for normal VGrid records.
        'Funding has many records and BestFit otherwise tends to compress them
        'until captions/values are heavily ellipsised.  On larger-font / high-DPI
        'interfaces scale this floor proportionally and allow horizontal scrolling
        'instead of squeezing all records into the viewport.
        Return Math.Max(
            120,
            CInt(Math.Ceiling(120.0F * Scale)))

    End Function

    Private Function GetMinimumVGridRowHeaderWidth(ByVal VG As VGridControl,
                                                   ByVal AppliedFont As Font) As Integer

        If VG Is Nothing Then Return 150

        Dim Scale As Single =
            GetVGridInterfaceScale(
                VG,
                AppliedFont)

        'The row-header panel contains the field captions and, for some VGrids,
        'embedded header editors.  Give it a slightly wider base minimum than a
        'record so bold/wrapped captions remain legible at high DPI.
        Return Math.Max(
            150,
            CInt(Math.Ceiling(150.0F * Scale)))

    End Function

    Private Function GetPreferredVGridContentWidth(
        ByVal VG As VGridControl,
        ByVal AvailableWidth As Integer) As Integer

        If VG Is Nothing Then Return Math.Max(1, AvailableWidth)

        Dim SafeAvailableWidth As Integer =
            Math.Max(1, AvailableWidth)

        Dim PreferredWidth As Integer =
            GetUncappedVGridContentWidth(VG)

        If PreferredWidth < 250 Then PreferredWidth = 250
        If PreferredWidth > SafeAvailableWidth Then PreferredWidth = SafeAvailableWidth

        Return PreferredWidth

    End Function

    Private Sub ApplyFontToControlTree(ByVal ParentControl As Control,
                                       ByVal NewFont As Font)

        If ParentControl Is Nothing OrElse ParentControl.IsDisposed Then Return

        ParentControl.Font = NewFont

        For Each Child As Control In ParentControl.Controls
            ApplyFontToControlTree(Child, NewFont)
        Next

    End Sub

    Private Function FindChildControls(Of T As Control)(ByVal ParentControl As Control) As List(Of T)

        Dim Result As New List(Of T)

        If ParentControl Is Nothing OrElse ParentControl.IsDisposed Then Return Result

        For Each Child As Control In ParentControl.Controls

            Dim Match As T = TryCast(Child, T)
            If Match IsNot Nothing Then Result.Add(Match)

            Result.AddRange(FindChildControls(Of T)(Child))

        Next

        Return Result

    End Function

    Public Sub InvalidateSectionFromDependency(ByVal SectionIndex As Integer,
                                               ByVal ChangedWorksheet As String)

        If Me.IsDisposed Then Return
        If SectionIndex < 0 OrElse SectionIndex >= DataPres.Sections.Length Then Return

        Dim State As InterfaceSectionRuntimeState = Nothing
        If Not SectionRuntimeStates.TryGetValue(SectionIndex, State) Then Return

        State.IsDirty = True
        State.NeedsPresentationRedefinition = True

        Dim IsSelectedSection As Boolean =
            XtraTabControlNewGIT.SelectedTabPageIndex = SectionIndex

        If Not AmActivated OrElse Not Me.Visible Then
            'A hidden/deactivated interface retains only metadata. Dispose any
            'stale controls now and recreate them only when the user returns.
            If State.IsBuilt Then UnloadSection(SectionIndex)
            Return
        End If

        If IsSelectedSection Then
            'IMPORTANT: do not dispose/rebuild the currently interacted-with VGrid
            'inside the same call stack that follows an Add Rows/Add Columns action.
            'DevExpress may still have queued MouseUp/layout/sorting preparation.
            QueueSectionRebuild(SectionIndex)
        ElseIf State.IsBuilt Then
            'An off-screen tab cannot be the control currently processing the mouse
            'action, so it is safe to unload it immediately and lazily recreate it.
            UnloadSection(SectionIndex)
        End If

    End Sub

    Private Sub QueueSectionRebuild(ByVal SectionIndex As Integer)

        If Me.IsDisposed Then Return
        If SectionIndex < 0 OrElse SectionIndex >= DataPres.Sections.Length Then Return

        PendingSectionRebuilds.Add(SectionIndex)

        If SectionRebuildTimer Is Nothing Then

            SectionRebuildTimer = New System.Windows.Forms.Timer With {
                .Interval = 75
            }

            AddHandler SectionRebuildTimer.Tick, AddressOf ProcessQueuedSectionRebuilds

        End If

        'Restarting coalesces several dependency notifications generated by the
        'same structural workbook operation into one rebuild pass.
        SectionRebuildTimer.Stop()
        SectionRebuildTimer.Start()

    End Sub

    Private Sub ProcessQueuedSectionRebuilds(ByVal sender As Object, ByVal e As EventArgs)

        If SectionRebuildTimer IsNot Nothing Then SectionRebuildTimer.Stop()
        If Me.IsDisposed Then Return

        Dim Pending As Integer() = PendingSectionRebuilds.ToArray()
        PendingSectionRebuilds.Clear()

        For Each SectionIndex As Integer In Pending

            Dim State As InterfaceSectionRuntimeState = Nothing
            If Not SectionRuntimeStates.TryGetValue(SectionIndex, State) Then Continue For
            If Not State.IsDirty Then Continue For

            If Not AmActivated OrElse Not Me.Visible Then
                If State.IsBuilt Then UnloadSection(SectionIndex)
                Continue For
            End If

            If XtraTabControlNewGIT.SelectedTabPageIndex = SectionIndex Then
                RebuildSelectedSectionAtomically(SectionIndex)
            ElseIf State.IsBuilt Then
                UnloadSection(SectionIndex)
            End If

        Next

    End Sub

    Private Sub UnloadSection(ByVal SectionIndex As Integer)

        If SectionIndex < 0 Then Return

        Dim VGridEditorHelpers As List(Of VGridRowInplaceEditorHelper) = Nothing

        If VGridInplaceEditorHelpersBySection.TryGetValue(SectionIndex, VGridEditorHelpers) Then

            For Each Helper As VGridRowInplaceEditorHelper In VGridEditorHelpers
                If Helper Is Nothing Then Continue For
                Try
                    Helper.DetachForDisposal()
                Catch
                End Try
            Next

            VGridInplaceEditorHelpersBySection.Remove(SectionIndex)

        End If

        Dim Extenders As List(Of VGridCategoryButtonExtender) = Nothing

        If VGridCategoryExtendersBySection.TryGetValue(SectionIndex, Extenders) Then

            For Each Extender As VGridCategoryButtonExtender In Extenders
                If Extender Is Nothing Then Continue For
                Try
                    Extender.RemoveCustomButton()
                Catch
                End Try
                VGridCategoryExtenders.Remove(Extender)
            Next

            VGridCategoryExtendersBySection.Remove(SectionIndex)

        End If

        If TPs IsNot Nothing AndAlso SectionIndex < TPs.Length AndAlso TPs(SectionIndex) IsNot Nothing Then

            Dim SectionDataSources As New List(Of AbovoUnboundSource)

            'Only detach our event handlers here.  Do NOT set VGrid.DataSource to
            'Nothing and do NOT clear VGrid.Rows: DevExpress must be allowed to
            'tear down its own ViewInfo/row hierarchy as part of Control.Dispose().
            DetachSectionControls(TPs(SectionIndex), SectionDataSources)

            Try
                TPs(SectionIndex).Dispose()
            Catch
            End Try

            'The owning controls are now disposed, so the unbound sources can be
            'disposed without forcing the VGrid through a live data-source reset.
            For Each UBS As AbovoUnboundSource In SectionDataSources.Distinct().ToArray()
                If UBS Is Nothing Then Continue For
                Try
                    UBS.Dispose()
                Catch
                End Try
            Next

            TPs(SectionIndex) = Nothing

        End If

        If AbovoTabPages IsNot Nothing AndAlso SectionIndex < AbovoTabPages.Length Then
            AbovoTabPages(SectionIndex) = New AbovoTabPage With {
                .Index = SectionIndex,
                .TabPage = XtraTabControlNewGIT.TabPages(SectionIndex)
            }
        End If

        Dim State As InterfaceSectionRuntimeState = Nothing
        If SectionRuntimeStates.TryGetValue(SectionIndex, State) Then State.IsBuilt = False

    End Sub

    Private Sub DetachSectionControls(ByVal ParentControl As Control,
                                      ByVal DataSourcesToDispose As List(Of AbovoUnboundSource))

        If ParentControl Is Nothing OrElse ParentControl.IsDisposed Then Return

        For Each Child As Control In ParentControl.Controls.Cast(Of Control).ToArray()

            DetachSectionControls(Child, DataSourcesToDispose)

            Dim GC As GridControl = TryCast(Child, GridControl)
            If GC IsNot Nothing Then

                RemoveHandler GC.ProcessGridKey, AddressOf GridControl_ProcessGridKey

                Dim UBS As AbovoUnboundSource = TryCast(GC.DataSource, AbovoUnboundSource)
                If UBS IsNot Nothing Then
                    RemoveHandler UBS.ValueNeeded, AddressOf UnboundDS_ValueNeeded
                    RemoveHandler UBS.ValuePushed, AddressOf UnboundDS_ValuePushed
                    If Not DataSourcesToDispose.Contains(UBS) Then DataSourcesToDispose.Add(UBS)
                End If

                'Do not break DataSource/view state manually. The GridControl owns
                'the view and will tear it down safely when its parent is disposed.
                Continue For

            End If

            Dim VG As VGridControl = TryCast(Child, VGridControl)
            If VG IsNot Nothing Then

                RemoveHandler VG.KeyDown, AddressOf VGrid_KeyDown
                RemoveHandler VG.EditorKeyDown, AddressOf VGrid_KeyDown
                RemoveHandler VG.CustomDrawRowValueCell, AddressOf VGrid_CustomDrawCell
                RemoveHandler VG.CustomDrawRowValueCell, AddressOf LiveVGrid_CustomDrawCell
                RemoveHandler VG.ValidatingEditor, AddressOf VGrid_ValidatingEditor
                RemoveHandler VG.ShowingEditor, AddressOf VGrid_ShowingEditor
                RemoveHandler VG.ShownEditor, AddressOf VGrid_ShownEditor
                RemoveHandler VG.DoubleClick, AddressOf VGrid_Event_DoubleClick
                RemoveHandler VG.CustomRecordCellEdit, AddressOf VGrid_CustomCellEditor
                RemoveHandler VG.CustomRecordCellEditForEditing, AddressOf VGrid_CellEditorForEditing

                Dim UBS As AbovoUnboundSource = TryCast(VG.DataSource, AbovoUnboundSource)
                If UBS IsNot Nothing Then
                    RemoveHandler UBS.ValueNeeded, AddressOf UnboundDS_ValueNeeded
                    RemoveHandler UBS.ValuePushed, AddressOf UnboundDS_ValuePushed
                    If Not DataSourcesToDispose.Contains(UBS) Then DataSourcesToDispose.Add(UBS)
                End If

                'Crucially, no VG.DataSource = Nothing and no VG.Rows.Clear() here.
                'Those operations can invalidate ViewInfo.Grid while DevExpress still
                'has queued mouse/layout work against the control.

            End If

        Next

    End Sub

    Sub GetSectionControlCollection(Section As PresentationSection, Parent As Object, SetSectionID As Integer, TP As TablePanel)

        Dim LastBottom As Integer = 0
        Dim ActiveDataSet As DataCellRange
        Dim TPRowCount As Integer = -1

        Dim SectionControlsCumlHeight As Integer = 0
        TP.BeginInit()

        'The XtraTabPage is the single outer scroll owner. The TablePanel must
        'therefore represent the FULL HEIGHT of its content. DockStyle.Fill
        'turns it into another viewport-sized surface and prevents AutoScroll
        'on the page from seeing content that extends below the visible area.
        TP.Dock = DockStyle.Top
        TP.AutoSize = True
        TP.AutoSizeMode = AutoSizeMode.GrowAndShrink

        Dim DefPaddin As New Padding

        DefPaddin.Left = CInt(Me.Width * 0.01)

        TP.Padding = DefaultTablePanelPadding

        If ParentGroupForm IsNot Nothing Then
            TP.Width = Math.Max(1, ParentGroupForm.ClientSize.Width - 40)
        Else
            TP.Width = Math.Max(1, Me.ClientSize.Width - 40)
        End If

        TP.BackColor = Color.White

        'BuildSection normally creates the standard four-column responsive
        'section layout before calling this routine.  Older/direct callers may
        'still supply an empty TablePanel, so initialise it only when required.
        '
        'Do NOT append another AutoSize column here.  The previous code did so
        'on every build, leaving five columns (four responsive columns plus one
        'full-width AutoSize column).  That made spanning controls resolve
        'differently at large resolutions.
        If TP.Columns.Count = 0 Then
            TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 2))
            TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 1))
            TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 5))
            TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Absolute, 100))
        End If

        'TODO AcContainers(AcContainersCount).Controls.Add(TP)

        For Each SectionElement In Section.SectionElements

            TPRowCount += 1

            TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 200))

            TP.Rows(TPRowCount).Tag = TPRowCount.ToString
#End Region

#Region "LiveGrid"

            If SectionElement.Type = "LiveGrid" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                'A Workings name may project non-contiguous worksheet areas (for
                'example the shared Year columns plus one of several result blocks).
                'Range.GetDataSource can bind only one contiguous range and a
                'worksheet cell may participate in only one such binding. Flattening
                'the projection into a range therefore made later tabs collide with
                'the first tab's binding. Use a detached read-only UnboundSource and
                'resolve each requested value from its original workbook cell.
                Dim LiveWorksheet As DevExpress.Spreadsheet.Worksheet =
                    ExcelModels(ModelID).WB.Worksheets(ActiveDataSet.SourceWorksheet)
                Dim LiveSourceRows As New List(Of Integer)
                Dim LiveSourceColumns As New List(Of Integer)
                Dim LiveSourceRanges As New List(Of DevExpress.Spreadsheet.CellRange)

                If ActiveDataSet.LiveGridSourceAreaReferences IsNot Nothing AndAlso
                   ActiveDataSet.LiveGridSourceAreaReferences.Count > 0 Then
                    For Each AreaReference As String In ActiveDataSet.LiveGridSourceAreaReferences
                        LiveSourceRanges.Add(LiveWorksheet.Range(AreaReference))
                    Next
                Else
                    LiveSourceRanges.Add(LiveWorksheet.Range(ActiveDataSet.DataRange))
                End If

                For RowIndex As Integer = LiveSourceRanges(0).TopRowIndex To LiveSourceRanges(0).BottomRowIndex
                    If LiveWorksheet.Rows(RowIndex).Visible Then LiveSourceRows.Add(RowIndex)
                Next

                For Each SourceRange As DevExpress.Spreadsheet.CellRange In LiveSourceRanges
                    For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex
                        If LiveWorksheet.Columns(ColumnIndex).Visible Then LiveSourceColumns.Add(ColumnIndex)
                    Next
                Next

                LiveSourceRows.RemoveAll(
                    Function(RowIndex) IsLiveProjectionRowBlank(
                        LiveWorksheet,
                        RowIndex,
                        LiveSourceColumns))

                UBSDataSourceCount += 1
                ReDim Preserve UnboundDataSources(UBSDataSourceCount)

                Dim LiveSetTag As New AbovoUnboundSourceTag With {
                    .ModelID = ModelID,
                    .GSID = GSID,
                    .CSID = CSID,
                    .RO = True,
                    .DSIndex = SectionElement.ControlSourceIndex,
                    .IsLiveGrid = True,
                    .LiveGridWorksheet = ActiveDataSet.SourceWorksheet,
                    .LiveGridSourceRows = LiveSourceRows,
                    .LiveGridSourceColumns = LiveSourceColumns
                }
                Dim LiveSource As New AbovoUnboundSource(UBSDataSourceCount, LiveSetTag)
                Dim LiveProperties As New List(Of UnboundSourceProperty)

                For ColumnOffset As Integer = 0 To ActiveDataSet.DataColumns.Length - 1
                    Dim LiveColumnName As String = "Col_" & ColumnOffset.ToString()
                    ActiveDataSet.DataColumns(ColumnOffset).ColumnTag.ActiveColumnName = LiveColumnName
                    LiveProperties.Add(New UnboundSourceProperty With {
                        .UserTag = ActiveDataSet.DataColumns(ColumnOffset).ColumnTag,
                        .DisplayName = ActiveDataSet.DataColumns(ColumnOffset).ColumnTag.ColumnHeading,
                        .Name = LiveColumnName,
                        .PropertyType = GetType(String)
                    })
                Next

                LiveSource.Properties.AddRange(LiveProperties)
                LiveSource.SetRowCount(LiveSourceRows.Count)
                AddHandler LiveSource.ValueNeeded, AddressOf UnboundDS_ValueNeeded
                UnboundDataSources(UBSDataSourceCount) = LiveSource


                GridCount += 1
                ReDim Preserve GridControls(GridCount)

                GridControls(GridCount) = New GridControl() With {
                    .Name = "GridControl_" & GridCount.ToString,
                    .Parent = Me,
                    .Dock = DockStyle.None,
                    .DataSource = LiveSource
                }

                CurrentAbovoTabPage.AddGrid(GridControls(GridCount))

                GridControls(GridCount).ForceInitialize()

                ' GridControls(GridCount).BeginUpdate()

                Formatter.FormatGridControl(GridControls(GridCount))

                GridViewCount += 1
                ReDim Preserve UsedGridVIEWS(GridViewCount)
                UsedGridVIEWS(GridViewCount) = New DevExpress.XtraGrid.Views.Grid.GridView

                UsedGridVIEWS(GridViewCount).Tag = New GridViewTag With {
                    .ModelID = ModelID,
                    .DSID = UBSDataSourceCount,
                    .DataSet = ActiveDataSet,
                    .IsLiveGrid = True,
                    .LiveGridWorksheet = ActiveDataSet.SourceWorksheet,
                    .LiveGridRange = ActiveDataSet.DataRange
                }

                UsedGridVIEWS(GridViewCount).ColumnPanelRowHeight += 2 * DefaultGridCellPadding

                UsedGridVIEWS(GridViewCount).UserCellPadding = New System.Windows.Forms.Padding(DefaultGridCellPadding)

                GridControls(GridCount).ViewCollection.Add(UsedGridVIEWS(GridViewCount))
                GridControls(GridCount).MainView = UsedGridVIEWS(GridViewCount)
                UsedGridVIEWS(GridViewCount).GridControl = GridControls(GridCount)
                ConfigureLiveGridView(UsedGridVIEWS(GridViewCount), ActiveDataSet)
                RegisterDataSetDependencies(SetSectionID, ActiveDataSet)
                ColCount = -1

                Dim LiveGridViewInfo As GridViewInfo = UsedGridVIEWS(GridViewCount).GetViewInfo()
                Dim IdealGridHeight As Integer =
                    (UsedGridVIEWS(GridViewCount).RowHeight *
                     LiveSourceRows.Count) +
                    LiveGridViewInfo.ColumnRowHeight +
                    DefaultTablePanelPadding.Top +
                    DefaultTablePanelPadding.Bottom
                Dim LiveGridHeightCap As Integer = MaxGridHeight
                If GSID = 2 Then
                    'Output statements need enough vertical space for compact
                    'summary reports. Use the current monitor's working area while
                    'retaining the established cap for assumptions and Workings.
                    LiveGridHeightCap = Math.Max(
                        MaxGridHeight,
                        Screen.FromControl(Me).WorkingArea.Height - 100)
                End If
                IdealGridHeight = Math.Min(IdealGridHeight, LiveGridHeightCap)
                GridControls(GridCount).Height = IdealGridHeight

                SectionControlsCumlHeight += IdealGridHeight + (2 * DefaultTablePanelPadding.Left)

                TP.Controls.Add(GridControls(GridCount))

                TP.SetCell(GridControls(GridCount), TPRowCount, 0)
                TP.SetColumnSpan(GridControls(GridCount), 4)

#End Region

#Region "LiveVGrid"

            ElseIf SectionElement.Type = "LiveVGrid" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                'Outputs cashflows are naturally read down as measures and across
                'as periods.  Project the workbook range through the same detached
                'source used by LiveGrid, then transpose only the presentation:
                'worksheet columns become VGrid rows and worksheet rows become
                'records.  No worksheet cell is data-bound or edited by this view.
                Dim LiveWorksheet As DevExpress.Spreadsheet.Worksheet =
                    ExcelModels(ModelID).WB.Worksheets(ActiveDataSet.SourceWorksheet)
                Dim LiveSourceRows As New List(Of Integer)
                Dim LiveSourceColumns As New List(Of Integer)
                Dim LiveSourceRanges As New List(Of DevExpress.Spreadsheet.CellRange)

                If ActiveDataSet.LiveGridSourceAreaReferences IsNot Nothing AndAlso
                   ActiveDataSet.LiveGridSourceAreaReferences.Count > 0 Then
                    For Each AreaReference As String In ActiveDataSet.LiveGridSourceAreaReferences
                        LiveSourceRanges.Add(LiveWorksheet.Range(AreaReference))
                    Next
                Else
                    LiveSourceRanges.Add(LiveWorksheet.Range(ActiveDataSet.DataRange))
                End If

                For RowIndex As Integer = LiveSourceRanges(0).TopRowIndex To LiveSourceRanges(0).BottomRowIndex
                    If LiveWorksheet.Rows(RowIndex).Visible Then LiveSourceRows.Add(RowIndex)
                Next

                For Each SourceRange As DevExpress.Spreadsheet.CellRange In LiveSourceRanges
                    For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex
                        If LiveWorksheet.Columns(ColumnIndex).Visible Then LiveSourceColumns.Add(ColumnIndex)
                    Next
                Next

                LiveSourceRows.RemoveAll(
                    Function(RowIndex) IsLiveProjectionRowBlank(
                        LiveWorksheet,
                        RowIndex,
                        LiveSourceColumns))

                UBSDataSourceCount += 1
                ReDim Preserve UnboundDataSources(UBSDataSourceCount)

                Dim LiveSetTag As New AbovoUnboundSourceTag With {
                    .ModelID = ModelID,
                    .GSID = GSID,
                    .CSID = CSID,
                    .RO = True,
                    .DSIndex = SectionElement.ControlSourceIndex,
                    .IsLiveGrid = True,
                    .LiveGridWorksheet = ActiveDataSet.SourceWorksheet,
                    .LiveGridSourceRows = LiveSourceRows,
                    .LiveGridSourceColumns = LiveSourceColumns
                }
                Dim LiveSource As New AbovoUnboundSource(UBSDataSourceCount, LiveSetTag)
                Dim LiveProperties As New List(Of UnboundSourceProperty)

                For ColumnOffset As Integer = 0 To ActiveDataSet.DataColumns.Length - 1
                    Dim LiveColumnName As String = "Col_" & ColumnOffset.ToString()
                    ActiveDataSet.DataColumns(ColumnOffset).ColumnTag.ActiveColumnName = LiveColumnName
                    LiveProperties.Add(New UnboundSourceProperty With {
                        .UserTag = ActiveDataSet.DataColumns(ColumnOffset).ColumnTag,
                        .DisplayName = ActiveDataSet.DataColumns(ColumnOffset).ColumnTag.ColumnHeading,
                        .Name = LiveColumnName,
                        .PropertyType = GetType(String)
                    })
                Next

                LiveSource.Properties.AddRange(LiveProperties)
                LiveSource.SetRowCount(LiveSourceRows.Count)
                AddHandler LiveSource.ValueNeeded, AddressOf UnboundDS_ValueNeeded
                UnboundDataSources(UBSDataSourceCount) = LiveSource

                VertGridCount += 1
                ReDim Preserve VertGridControls(VertGridCount)
                Dim LiveVGrid As New VGridControl() With {
                    .Name = "LiveVGridControl_" & VertGridCount.ToString(),
                    .Parent = Me,
                    .Dock = DockStyle.None,
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Left,
                    .DataSource = LiveSource,
                    .LayoutStyle = LayoutViewStyle.MultiRecordView,
                    .ScrollVisibility = DevExpress.XtraVerticalGrid.ScrollVisibility.Horizontal,
                    .RecordWidth = 100,
                    .RowHeaderWidth = 340,
                    .RecordHeaderHeight = 42
                }
                VertGridControls(VertGridCount) = LiveVGrid
                CurrentAbovoTabPage.AddVGrid(LiveVGrid)
                LiveVGrid.ForceInitialize()

                LiveVGrid.OptionsBehavior.Editable = False
                LiveVGrid.OptionsBehavior.CopyToClipboardWithRowHeaders = True
                LiveVGrid.OptionsSelectionAndFocus.MultiSelect = True
                LiveVGrid.OptionsSelectionAndFocus.MultiSelectMode = MultiSelectMode.CellSelect
                LiveVGrid.OptionsView.ShowRecordHeaders = True

                Dim RecordHeaderColumnCount As Integer = 2
                Integer.TryParse(ActiveDataSet.LiveVGridRecordHeaderColumns, RecordHeaderColumnCount)
                RecordHeaderColumnCount = Math.Max(1, Math.Min(RecordHeaderColumnCount, LiveSourceColumns.Count))
                LiveVGrid.RecordHeaderFormat = "{Col_" & (RecordHeaderColumnCount - 1).ToString() & "}"

                Dim CategoryRowIndex As Integer = -1
                Dim ParsedCategoryRow As Integer
                If Integer.TryParse(ActiveDataSet.LiveVGridCategoryRow, ParsedCategoryRow) AndAlso
                   ParsedCategoryRow > 0 Then
                    CategoryRowIndex = ParsedCategoryRow - 1
                End If

                LiveVGrid.Rows.Clear()
                Dim CurrentCategory As CategoryRow = Nothing
                Dim LastCategoryCaption As String = Nothing

                For ColumnOffset As Integer = 0 To Math.Min(LiveSourceColumns.Count, ActiveDataSet.DataColumns.Length) - 1
                    Dim SourceColumnIndex As Integer = LiveSourceColumns(ColumnOffset)
                    Dim WorkbookRowCaption As String = BuildLiveGridColumnCaption(
                        LiveWorksheet,
                        LiveSourceRanges(0),
                        SourceColumnIndex,
                        String.Empty,
                        ActiveDataSet.LiveGridHeaderRows)
                    Dim RowCaption As String = WorkbookRowCaption
                    If String.IsNullOrWhiteSpace(RowCaption) Then
                        RowCaption = ActiveDataSet.DataColumns(ColumnOffset).ColumnTag.ColumnHeading
                    End If

                    Dim VRow As New EditorRow("Col_" & ColumnOffset.ToString()) With {
                        .Height = IdealGridRowHeight,
                        .Tag = New LiveVGridRowTag With {.SourceColumnIndex = SourceColumnIndex}
                    }
                    VRow.Properties.Caption = RowCaption
                    VRow.Properties.ReadOnly = True

                    If ColumnOffset < RecordHeaderColumnCount Then
                        LiveVGrid.Rows.Add(VRow)
                        'Record-header fields belong across the top of a VGrid, not
                        'as duplicate value rows. Set Visible only after insertion;
                        'DevExpress resets it when a detached row is first attached.
                        VRow.Visible = False
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(WorkbookRowCaption) AndAlso
                       IsLiveProjectionColumnBlank(
                           LiveWorksheet,
                           SourceColumnIndex,
                           LiveSourceRows) Then
                        LiveVGrid.Rows.Add(VRow)
                        VRow.Visible = False
                        Continue For
                    End If

                    Dim CategoryCaption As String = Nothing
                    If CategoryRowIndex >= 0 Then
                        CategoryCaption = LiveWorksheet.Cells(CategoryRowIndex, SourceColumnIndex).DisplayText.Trim()
                    End If

                    If Not String.IsNullOrWhiteSpace(CategoryCaption) AndAlso
                       Not String.Equals(CategoryCaption, LastCategoryCaption, StringComparison.Ordinal) Then
                        CurrentCategory = New CategoryRow("LiveCategory_" & ColumnOffset.ToString()) With {
                            .Height = 36,
                            .Tag = New LiveVGridCategoryTag With {
                                .SourceRowIndex = CategoryRowIndex,
                                .SourceColumnIndex = SourceColumnIndex
                            }
                        }
                        CurrentCategory.Properties.Caption = CategoryCaption
                        LiveVGrid.Rows.Add(CurrentCategory)
                        LastCategoryCaption = CategoryCaption
                    End If

                    If CurrentCategory Is Nothing Then
                        LiveVGrid.Rows.Add(VRow)
                    Else
                        CurrentCategory.ChildRows.Add(VRow)
                    End If
                Next

                LiveVGrid.Tag = New VGridLayoutTag With {
                    .TableRowIndex = TPRowCount,
                    .IsLiveGrid = True
                }
                AddHandler LiveVGrid.CustomDrawRowValueCell, AddressOf LiveVGrid_CustomDrawCell
                RegisterDataSetDependencies(SetSectionID, ActiveDataSet)

                Dim IdealVGridHeight As Integer = Math.Min(
                    Math.Max(400, GetPreferredVGridContentHeight(LiveVGrid) + 50),
                    MaxGridHeight)
                LiveVGrid.Height = IdealVGridHeight
                SectionControlsCumlHeight += IdealVGridHeight +
                    DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom
                TP.Controls.Add(LiveVGrid)
                TP.SetCell(LiveVGrid, TPRowCount, 0)
                TP.SetColumnSpan(LiveVGrid, 4)
                TP.AutoSize = True
                TP.AutoSizeMode = AutoSizeMode.GrowAndShrink
                UnboundDataSources(UBSDataSourceCount).AttachedVertGrid = LiveVGrid

#End Region

#Region "Grid"

            ElseIf SectionElement.Type = "Grid" Then

                Dim ApplyBands As Boolean = DataPres.DataSets(SectionElement.ControlSourceIndex).HasBands
                Dim HasActions As Boolean = False
                Dim HasLockRules As Boolean = False
                Dim BandColumButtonsAndDrawing As CombinedColumnBandButton_HeaderFooterDrawer = Nothing
                Dim CustSummariesAddedToGrid As Boolean = False
                Dim BandedFooterDone As Boolean = False

                ColCount = -1

                FooterOn = False
                FooterDone = False

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)


                ColList = New List(Of String)

                PropertiesCount = -1 'reset
                UBSDataSourceCount += 1
                ReDim Preserve UnboundDataSources(UBSDataSourceCount)
                Dim SetTag As New AbovoUnboundSourceTag With {.GSID = GSID, .CSID = CSID, .RO = ActiveDataSet.RO, .DSIndex = SectionElement.ControlSourceIndex}

                UnboundDataSources(UBSDataSourceCount) = New AbovoUnboundSource(UBSDataSourceCount, SetTag)

                RegisterDataSetDependencies(SetSectionID, ActiveDataSet)


                Dim IdealGridWidth As Double = 0

                For Each PresentedColumn In ActiveDataSet.DataColumns

                    If PresentedColumn.ColumnTag.IsCalculated Then UnboundDataSources(UBSDataSourceCount).UBSTag.HasCalcs = True

                    ColCount += 1
                    ColName = "Col_" & ColCount.ToString
                    PresentedColumn.ColumnTag.ActiveColumnName = ColName
                    PropertiesCount += 1
                    ReDim Preserve PropertyArray(PropertiesCount)

                    Select Case PresentedColumn.ColumnTag.DataType

                        Case "S"
                            PropType = GetType(String)

                        Case "I", "Y", "D"
                            PropType = GetType(Integer)

                        Case "N", "P", "C", "M", "SM", "R"
                            PropType = GetType(Double)

                        Case "B"
                            PropType = GetType(Integer)

                        Case Else
                            PropType = GetType(String)

                    End Select


                    PropertyArray(PropertiesCount) = New UnboundSourceProperty With {
                        .UserTag = PresentedColumn.ColumnTag,
                        .DisplayName = " " & PresentedColumn.ColumnTag.ColumnHeading & " ",
                        .Name = ColName,
                        .PropertyType = PropType
                    }

                    ColList.Add(ColName)

                Next

                PropertyList = PropertyArray

                UnboundDataSources(UBSDataSourceCount).Properties.AddRange(PropertyList)

                'Add dataaccess/push handlers
                AddHandler UnboundDataSources(UBSDataSourceCount).ValueNeeded, AddressOf UnboundDS_ValueNeeded
                AddHandler UnboundDataSources(UBSDataSourceCount).ValuePushed, AddressOf UnboundDS_ValuePushed

                'Create new grid for DS

                GridCount += 1
                ReDim Preserve GridControls(GridCount)


                GridControls(GridCount) = New GridControl() With {
                    .Name = "GridControl_" & GridCount.ToString,
                    .Parent = Me,
                    .Dock = DockStyle.Top,
                    .DataSource = UnboundDataSources(UBSDataSourceCount)
                }

                GridControls(GridCount).ForceInitialize()

                CurrentAbovoTabPage.AddGrid(GridControls(GridCount))
                AddHandler GridControls(GridCount).ProcessGridKey, AddressOf GridControl_ProcessGridKey

                UnboundDataSources(UBSDataSourceCount).AttachedGrid = GridControls(GridCount)

                GridControls(GridCount).BeginUpdate()


                GridViewCount += 1
                ReDim Preserve UsedGridVIEWS(GridViewCount)

                UsedGridVIEWS(GridViewCount) = New DevExpress.XtraGrid.Views.Grid.GridView

                UsedGridVIEWS(GridViewCount).BeginUpdate()

                UsedGridVIEWS(GridViewCount).Tag = New GridViewTag With {.ModelID = ModelID, .DSID = UBSDataSourceCount, .DataSet = ActiveDataSet}

                If ActiveDataSet.RO Then

                    UsedGridVIEWS(GridViewCount).OptionsBehavior.Editable = False
                    UsedGridVIEWS(GridViewCount).OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.None

                Else

                    UsedGridVIEWS(GridViewCount).OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.Update

                End If

                GridControls(GridCount).RepositoryItems.Add(CustCalcEditInteger)
                GridControls(GridCount).RepositoryItems.Add(CustCalcEditDecimal)

                If ActiveDataSet.HasValidations Then


                    For Each ValList In ActiveDataSet.ValidationLists

                        GridControls(GridCount).RepositoryItems.Add(RepositaryItems.GetEditorFromList(ValList).RetCombo)

                    Next

                End If

                UsedGridVIEWS(GridViewCount).AccessibleName = "GridView_" & GridViewCount.ToString
                UsedGridVIEWS(GridViewCount).Name = "GridView_" & GridViewCount.ToString
                GridControls(GridCount).ViewCollection.Add(UsedGridVIEWS(GridViewCount))
                GridControls(GridCount).MainView = UsedGridVIEWS(GridViewCount)
                UsedGridVIEWS(GridViewCount).GridControl = GridControls(GridCount)


                Dim FillSize As Integer = ActiveDataSet.RowCount
                UnboundDataSources(UBSDataSourceCount).SetRowCount(FillSize)


                For Each ColCheck As UnboundSourceProperty In PropertyArray

                    With UsedGridVIEWS(GridViewCount).Columns(ColCheck.UserTag.ActiveColumnName)

                        .Tag = ColCheck
                        .AppearanceCell.options.usebackcolor = True
                        .AppearanceCell.Options.UseBorderColor = True
                        .AppearanceCell.BorderColor = System.Drawing.Color.Silver
                        .AppearanceCell.backcolor = System.Drawing.Color.Silver
                        .AppearanceCell.forecolor = System.Drawing.Color.Black

                    End With

                    If ColCheck.UserTag.IsReadOnly OrElse ColCheck.UserTag.IsCalculated Then

                        With UsedGridVIEWS(GridViewCount).Columns(ColCheck.UserTag.ActiveColumnName)

                            .OptionsColumn.TabStop = False
                            .OptionsColumn.ReadOnly = True
                            .AppearanceCell.BorderColor = System.Drawing.Color.White
                            .appearanceCell.backColor = System.Drawing.Color.White
                            .AppearanceCell.ForeColor = System.Drawing.Color.DarkGray

                        End With

                    End If

                    If ColCheck.UserTag.IsFixed Then

                        UsedGridVIEWS(GridViewCount).OptionsView.ColumnAutoWidth = False
                        UsedGridVIEWS(GridViewCount).Columns(ColCheck.UserTag.ActiveColumnName).Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left

                    End If

                Next



                GridControls(GridCount).Top = LastBottom + 20

                Dim ThisCol As Integer

                For ThisCol = 0 To UsedGridVIEWS(GridViewCount).Columns.Count - 1

                    UsedGridVIEWS(GridViewCount).Columns(ThisCol).Tag = ActiveDataSet.DataColumns(ThisCol).ColumnTag

                Next

                Dim ColTag As DataColumnTag
                Dim GVcolumn As GridColumn
                Dim siTotal As GridColumnSummaryItem

                Dim DefWidth As Integer

                'Add the "Insert Rows" button


                For Each GVcolumn In UsedGridVIEWS(GridViewCount).Columns

                    GVcolumn.Width += 2 * DefaultGridCellPadding
                    ColTag = GVcolumn.Tag
                    DefWidth = GVcolumn.Width

                    With GVcolumn

                        .Caption = " " & ColTag.ColumnHeading & " "
                        .ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowForFocusedCell
                        .AppearanceHeader.BackColor = System.Drawing.Color.White
                        .AppearanceHeader.Options.UseBackColor = True
                        .AppearanceHeader.ForeColor = AbovoBlue
                        .AppearanceHeader.Options.UseForeColor = True
                        .AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom
                        .AppearanceHeader.BorderColor = System.Drawing.Color.White
                        .AppearanceHeader.Options.UseBorderColor = True

                    End With

                    Dim EditControl As RepositaryItems.AbovoRespositaryItem

                    ColTag.DefaultColumnWidth = DefWidth

                    If ColTag.IsDummyColumn Then

                        GVcolumn.OptionsColumn.ReadOnly = True

                        GVcolumn.MinWidth = 50
                        GVcolumn.MaxWidth = 50
                        GVcolumn.Width = 50
                        GVcolumn.AppearanceCell.BackColor = Color.White
                        GVcolumn.AppearanceCell.ForeColor = Color.White
                        GVcolumn.AppearanceHeader.ForeColor = Color.White
                        GVcolumn.AppearanceHeader.ForeColor = Color.White

                    End If

                    If ColTag.IsReadOnly OrElse ColTag.IsCalculated Then

                        GVcolumn.AppearanceCell.BackColor = Color.WhiteSmoke
                        GVcolumn.AppearanceCell.ForeColor = AbovoBlue

                    End If

                    Dim UseCombo As Boolean = False

                    If Not IsNothing(ColTag.RepositaryID) Then

                        ColTag.HasComboEdit = True
                        UseCombo = True
                        EditControl = RepositaryItems.GetEditor(ColTag.RepositaryID, ModelID)

                        If EditControl.RepType = "CMB" Then

                            GridCombosCount += 1
                            EditControl.RetCombo.Name = "Combo_" & GridCombosCount.ToString
                            AddHandler EditControl.RetCombo.Enter, AddressOf ComboOpen
                            GridControls(GridCount).RepositoryItems.Add(EditControl.RetCombo)
                            ComboReposClassesCount += 1

                            ReDim Preserve ComboReposClasses(ComboReposClassesCount)

                            ComboReposClasses(ComboReposClassesCount) = New AbovoGridRespoitaryCombo With {
                                .ID = ComboReposClassesCount,
                                .RepoistaryID = ColTag.RepositaryID,
                                .Combo = EditControl.RetCombo,
                                .ModelID = ModelID,
                                .GridID = GridCount}
                            CurrentAbovoTabPage.AddRepCombo(ComboReposClasses(ComboReposClassesCount))
                            GVcolumn.ColumnEdit = EditControl.RetCombo

                        End If

                    End If

                    If Not IsNothing(ColTag.TipText) Then

                        GVcolumn.ToolTip = TrimSpaces(ColTag.TipText) 'HelpManager.CreateSuperTooltip(ColTag.ColumnHeading, ColTag.TipText, "1")

                    End If

                    Select Case ColTag.DataType

                        Case "S"

                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.None
                            ColTag.ShowDefaultmask = -1

                        Case "M"

                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "c0"
                            Dim edit As New RepositoryItemTextEdit()
                            edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                            edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                            edit.Mask.EditMask = "c5"
                            GVcolumn.ColumnEdit = edit
                            ColTag.DefaultTextEditor = edit
                            ColTag.ShowDefaultmask = 0

                        Case "SM"

                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "c2"
                            Dim edit As New RepositoryItemTextEdit()
                            edit.UseMaskAsDisplayFormat = True
                            edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                            edit.Mask.EditMask = "c2"
                            GVcolumn.ColumnEdit = edit
                            ColTag.DefaultTextEditor = edit
                            ColTag.ShowDefaultmask = 1

                        Case "I"

                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "#,##0"
                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                            ColTag.ShowDefaultmask = -1

                            If Not UseCombo Then

                                Dim edit As New RepositoryItemTextEdit()
                                edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                                edit.Mask.EditMask = "n0"
                                GVcolumn.ColumnEdit = edit
                                ColTag.DefaultTextEditor = edit

                            End If

                        Case "C"

                            'GVcolumn.ColumnEdit = New RepositoryItemTextEdit With {
                            '        .AllowNullInput = True
                            '        }

                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "#,###,##0" '{#,###,##0;(#,###,##0)}"
                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

                        Case "R"

                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "#,##0.00"
                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

                            Dim edit As New RepositoryItemTextEdit()
                            edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                            edit.Mask.EditMask = "n5"
                            GVcolumn.ColumnEdit = edit
                            ColTag.DefaultTextEditor = edit

                        Case "P"

                            GVcolumn.Width += 2 * DefaultGridCellPadding
                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "p2"

                            ColTag.ShowDefaultmask = 3
                            Dim SetMinVal As Object = IIf(ColTag.MinVal Is Nothing, 0, IIf(ColTag.MinVal = "NOMIN", Nothing, CDec(ColTag.MinVal)))
                            Dim SetMaxVal As Object = IIf(ColTag.MaxVal Is Nothing, 1, IIf(ColTag.MaxVal = "NOMAX", Nothing, CDec(ColTag.MaxVal)))

                            If Not ColTag.IsReadOnly Then

                                Dim edit As New RepositoryItemSpinEdit()

                                edit.MinValue = IIf(SetMinVal = Nothing, Nothing, CDec(SetMinVal))
                                edit.Increment = IIf(ColTag.DefIncrement Is Nothing, CDec(0.0025), CDec(ColTag.DefIncrement))

                                edit.MaxValue = IIf(SetMaxVal = Nothing, Nothing, CDec(SetMaxVal))
                                edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                                edit.Mask.EditMask = "p2"
                                edit.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                                edit.DisplayFormat.FormatString = "p2"
                                GVcolumn.ColumnEdit = edit
                                ColTag.DefaultTextEditor = edit

                            End If

                        Case "B"

                        Case Else



                    End Select

                    If ColTag.ShowSummary = "Sum" Then

                        If Not FooterDone Then SetFooterOn(UsedGridVIEWS(GridViewCount))

                        siTotal = New GridColumnSummaryItem
                        siTotal.SummaryType = DevExpress.Data.SummaryItemType.Sum
                        siTotal.DisplayFormat = "{0:n0}"

                        GVcolumn.Summary.Add(siTotal)

                    ElseIf ColTag.ShowSummary = "Count" Then

                        If Not FooterDone Then SetFooterOn(UsedGridVIEWS(GridViewCount))

                        siTotal = New GridColumnSummaryItem
                        siTotal.SummaryType = DevExpress.Data.SummaryItemType.Count
                        siTotal.DisplayFormat = "{0:n0}"
                        GVcolumn.Summary.Add(siTotal)

                    End If

                    If ColTag.HasActions Then

                        If HasActions = False Then

                            HasActions = True

                        End If

                        UsedGridVIEWS(GridViewCount).InvalidateColumnHeader(GVcolumn)

                    End If

                    IdealGridWidth += GVcolumn.Width

                Next

                With UsedGridVIEWS(GridViewCount)

                    .OptionsSelection.MultiSelect = True
                    .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect

                    .OptionsBehavior.FocusLeaveOnTab = True
                    .OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDownFocused

                    .OptionsView.RowAutoHeight = True
                    .OptionsView.BestFitMaxRowCount = -1
                    '.OptionsView.ColumnAutoWidth = True
                    .OptionsView.BestFitMaxRowCount = ActiveDataSet.UsedRows

                    .ColumnPanelRowHeight += 4 * DefaultGridCellPadding
                    .UserCellPadding = New System.Windows.Forms.Padding(DefaultGridCellPadding)

                End With

                If UnboundDataSources(UBSDataSourceCount).UBSTag.RO Then UsedGridVIEWS(GridViewCount).OptionsBehavior.Editable = False

                Formatter.FormatGridView(UsedGridVIEWS(GridViewCount), GridControls(GridCount))

                LastBottom = GridControls(GridCount).Bottom
                UsedGridVIEWS(GridViewCount).EndUpdate()
                GridControls(GridCount).Height += 100

#Region "Bands"

                If ApplyBands Then




                    Dim LastBand As String = ""
                    Dim LastBandTitle As String = "StartingBandDummyName"
                    Dim OriginColTag As DataColumnTag = Nothing
                    Dim LastColour As Color = Color.Red
                    BandGridViewsCount += 1
                    ReDim Preserve UsedBANDedGridVIEWS(BandGridViewsCount)

                    UsedBANDedGridVIEWS(BandGridViewsCount) = New BandedGridView With {
                        .Tag = UsedGridVIEWS(GridViewCount).Tag,
                        .UserCellPadding = New System.Windows.Forms.Padding(DefaultGridCellPadding),
                        .HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Auto,
                        .VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Auto
                    }

                    BandColumButtonsAndDrawing = New CombinedColumnBandButton_HeaderFooterDrawer(UsedBANDedGridVIEWS(BandGridViewsCount), Me, SetSectionID)

                    UsedBANDedGridVIEWS(BandGridViewsCount).BeginUpdate()



                    With UsedBANDedGridVIEWS(BandGridViewsCount)

                        .OptionsSelection.MultiSelect = True
                        .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect

                        .OptionsBehavior.FocusLeaveOnTab = True
                        .OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDownFocused

                        .OptionsView.RowAutoHeight = True
                        '.OptionsView.ColumnAutoWidth = True
                        .OptionsView.BestFitMaxRowCount = ActiveDataSet.UsedRows

                        .ColumnPanelRowHeight += 4 * DefaultGridCellPadding
                        .UserCellPadding = New System.Windows.Forms.Padding(DefaultGridCellPadding)

                        .OptionsCustomization.AllowChangeColumnParent = False
                        .BandPanelRowHeight += (4 * DefaultGridCellPadding)

                    End With

                    If ActiveDataSet.RO Then

                        UsedBANDedGridVIEWS(BandGridViewsCount).OptionsBehavior.Editable = False
                        UsedBANDedGridVIEWS(BandGridViewsCount).OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.None

                    Else

                        UsedBANDedGridVIEWS(BandGridViewsCount).OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.Update

                    End If

                    IdealGridWidth = 0

                    'Dim CurrBand As New GridBand

                    For Each GVcolumn In UsedGridVIEWS(GridViewCount).Columns

                        OriginColTag = GVcolumn.Tag

                        If OriginColTag.BandID <> LastBandTitle Then

                            BandGridViewBandsCount += 1
                            ReDim Preserve UsedBANDedGridViewBANDS(BandGridViewBandsCount)

                            Dim BandCaption As String = IIf(OriginColTag.BandID Is Nothing, " ", OriginColTag.BandID)

                            BandCaption = BandCaption.Replace("vblf", "<br>")

                            If OriginColTag.BandID Is Nothing Then

                                BandCaption = " "

                            End If

                            UsedBANDedGridViewBANDS(BandGridViewBandsCount) = New GridBand With {.Caption = BandCaption}

                            If Not IsNothing(OriginColTag.BandTipText) Then UsedBANDedGridViewBANDS(BandGridViewBandsCount).ToolTip = TrimSpaces(OriginColTag.BandTipText)

                            If OriginColTag.BandID = "FixedLeft" Then

                                UsedBANDedGridViewBANDS(BandGridViewBandsCount).Caption = " "
                                UsedBANDedGridViewBANDS(BandGridViewBandsCount).Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left

                            End If


                            With UsedBANDedGridViewBANDS(BandGridViewBandsCount)

                                .AppearanceHeader.ForeColor = AbovoBlue
                                .AppearanceHeader.Options.UseForeColor = True
                                .AppearanceHeader.BackColor = Color.White
                                .AppearanceHeader.Options.UseBackColor = True
                                .AppearanceHeader.Options.UseTextOptions = True
                                .AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
                                .AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom

                            End With

                            If LastColour = Color.Red Then

                                LastColour = Color.White

                            ElseIf LastColour = Color.White Then

                                LastColour = Color.Wheat

                            ElseIf LastColour = Color.Wheat Then

                                LastColour = AbovoBlue

                            ElseIf LastColour = AbovoBlue Then

                                LastColour = Color.Wheat

                            End If

                            Dim EditText As String


                            If OriginColTag.BandEditDescription IsNot Nothing Then

                                EditText = OriginColTag.BandEditDescription

                            Else

                                EditText = "Edit " & GVcolumn.FieldName & " of " & OriginColTag.BandID

                            End If

                            Dim ActToken As New ActionToken With {
                                    .ActionType = OriginColTag.EditRepNRHereExpansionMethod,
                                    .ActionStrData1 = OriginColTag.ActionNR,
                                    .ActionNR = OriginColTag.ActionNR,
                                    .ActionNumber1 = SetSectionID,
                                    .ActionStrData2 = EditText} ', .ActionData3 = ActiveDataSet.RepeatingHeaderText}

                            BandColumButtonsAndDrawing.AddCustomBandAddColumsButton(UsedBANDedGridViewBANDS(BandGridViewBandsCount), ActToken)

                            '           .HasActions = OriginColTag.HasActions,

                            Dim NewBandTag As New BandTag With {
                                    .ID = BandGridViewsCount,
                                    .ActionToken = ActToken,
                                    .ActionNR = OriginColTag.ActionNR,
                                    .ActionDescription = OriginColTag.BandEditDescription,
                                    .HighLightColour = LastColour}

                            If OriginColTag.BandID IsNot Nothing Then NewBandTag.DoBorder = True

                            UsedBANDedGridViewBANDS(BandGridViewBandsCount).Tag = NewBandTag

                            UsedBANDedGridVIEWS(BandGridViewsCount).Bands.Add(UsedBANDedGridViewBANDS(BandGridViewBandsCount))

                            LastBandTitle = OriginColTag.BandID

                        End If

                        BandGridViewColsCount += 1
                        ReDim Preserve UsedBANDedGridViewCOLS(BandGridViewColsCount)

                        If OriginColTag.HasActions Then

                            UsedBANDedGridViewBANDS(BandGridViewBandsCount).Tag.HasActions = True
                            UsedBANDedGridViewBANDS(BandGridViewBandsCount).Tag.ActionNR = OriginColTag.ActionNR

                        End If

                        UsedBANDedGridViewCOLS(BandGridViewColsCount) = New BandedGridColumn With {
                             .Tag = OriginColTag,
                            .FieldName = GVcolumn.FieldName,
                            .Width = GVcolumn.Width,
                            .Visible = True,
                            .ColumnEdit = GVcolumn.ColumnEdit,
                            .Fixed = GVcolumn.Fixed,
                            .Caption = GVcolumn.Caption,
                            .SortIndex = GVcolumn.SortIndex
                            }

                        If GVcolumn.Fixed Then

                            UsedBANDedGridViewCOLS(BandGridViewColsCount).Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left

                        End If
                        UsedBANDedGridViewCOLS(BandGridViewColsCount).AppearanceCell.TextOptions.HAlignment = GVcolumn.AppearanceCell.TextOptions.HAlignment

                        With UsedBANDedGridViewCOLS(BandGridViewColsCount)

                            .AppearanceCell.Options.UseBackColor = True
                            .AppearanceCell.Options.UseBorderColor = True
                            .AppearanceCell.BorderColor = GVcolumn.AppearanceCell.BorderColor
                            .AppearanceCell.BackColor = GVcolumn.AppearanceCell.BackColor

                        End With

                        With UsedBANDedGridViewCOLS(BandGridViewColsCount)

                            .AppearanceHeader.ForeColor = AbovoBlue
                            .AppearanceHeader.Options.UseForeColor = True
                            .AppearanceHeader.BackColor = Color.White
                            .AppearanceHeader.Options.UseBackColor = True

                            .AppearanceHeader.TextOptions.HAlignment = GVcolumn.AppearanceHeader.TextOptions.HAlignment
                            .AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom
                            .AppearanceHeader.Options.UseTextOptions = True

                            .DisplayFormat.FormatType = GVcolumn.DisplayFormat.FormatType
                            .DisplayFormat.FormatString = GVcolumn.DisplayFormat.FormatString

                            .ToolTip = GVcolumn.ToolTip
                            .OptionsColumn.ReadOnly = GVcolumn.ReadOnly

                        End With

                        If OriginColTag.ShowSummary = "Sum" Then

                            If Not BandedFooterDone Then

                                BandedFooterDone = True
                                UsedBANDedGridVIEWS(BandGridViewsCount).OptionsView.ShowFooter = True
                                UsedBANDedGridVIEWS(BandGridViewsCount).FooterPanelHeight += (2 * DefaultGridCellPadding)

                            End If

                            siTotal = New GridColumnSummaryItem
                            siTotal.SummaryType = DevExpress.Data.SummaryItemType.Sum
                            siTotal.DisplayFormat = "{0:n0}"

                            UsedBANDedGridViewCOLS(BandGridViewColsCount).Summary.Add(siTotal)

                        ElseIf OriginColTag.ShowSummary = "Count" Then

                            If Not FooterDone Then SetFooterOn(UsedGridVIEWS(GridViewCount))

                            siTotal = New GridColumnSummaryItem
                            siTotal.SummaryType = DevExpress.Data.SummaryItemType.Count
                            siTotal.DisplayFormat = "{0:n0}"
                            UsedBANDedGridViewCOLS(BandGridViewColsCount).Summary.Add(siTotal)

                            If Not BandedFooterDone Then

                                BandedFooterDone = True
                                UsedBANDedGridVIEWS(BandGridViewsCount).OptionsView.ShowFooter = True
                                UsedBANDedGridVIEWS(BandGridViewsCount).FooterPanelHeight += (2 * DefaultGridCellPadding)

                            End If

                        End If

                        UsedBANDedGridVIEWS(BandGridViewsCount).Columns.Add(UsedBANDedGridViewCOLS(BandGridViewColsCount))
                        UsedBANDedGridViewBANDS(BandGridViewBandsCount).Columns.Add(UsedBANDedGridViewCOLS(BandGridViewColsCount))

                        IdealGridWidth += GVcolumn.Width

                    Next

                    LastBandTitle = "LastBandDummyName"

                    Dim EditControlInitialised As Boolean = False
                    Dim InEditorAddingMode As Boolean = False
                    Dim DoneFirstItem As Boolean = False
                    Dim MaxItemLength As Integer = 5
                    Dim SetIndex As Integer = -1
                    Dim SetColIndex As Integer = -1
                    Dim CurrBaseCaption As String = ""

                    Formatter.FormatGridView(UsedBANDedGridVIEWS(BandGridViewsCount), GridControls(GridCount))

                    For Each BGC As BandedGridColumn In UsedBANDedGridVIEWS(BandGridViewsCount).Columns

                        OriginColTag = BGC.Tag

                        If OriginColTag.BandID <> LastBandTitle Then

                            SetIndex += 1
                            SetColIndex = 0

                            EditControlInitialised = False
                            InEditorAddingMode = False

                            If OriginColTag.EditRepNRHere Then
                                InEditorAddingMode = True
                            Else
                                InEditorAddingMode = False
                            End If

                            DoneFirstItem = False
                            MaxItemLength = 5

                        End If

                        If InEditorAddingMode Then

                            If OriginColTag.EditRepNRHereROColumn Then

                                'With BGC

                                '    .AppearanceHeader.Options.UseBackColor = True
                                '    .AppearanceHeader.BackColor = AbovoComboBGC
                                '    .AppearanceHeader.Options.UseForeColor = True
                                '    .AppearanceHeader.ForeColor = Color.White 'AbovoComboBGC

                                'End With

                            ElseIf OriginColTag.BandID Is Nothing Then

                                'do nothing - read only columns cannot have in-header editors

                            Else

                                If UCase(OriginColTag.EditRepNRHereEditor) = "COMBO" Then

                                    InHeaderCombosCount += 1
                                    ReDim Preserve InHeaderCombos(InHeaderCombosCount)

                                    Dim EditControl As New AbovoDEHeaderComboBox
                                    EditControl.AddBlankFirstItem = True
                                    EditControl.InitialiseStandard(OriginColTag.EditRepNRHereComboRepository)
                                    EditControl.EditValue = OriginColTag.EditRepNRHereInitialValue

                                    Dim RetComb As RepositoryItemComboBox = EditControl.Properties

                                    Dim InColumnEditorTag As New InColumnEditorTagCombo With {
                                        .EditingNRIndexPosition = OriginColTag.EditNRIndexPosition,
                                        .EditingNRName = OriginColTag.RepeatingNR,
                                        .NROrientation = IIf(OriginColTag.EditRepNRNROrientation = "PORT", Orientation.Vertical, Orientation.Horizontal),
                                        .EditorType = "COMBO",
                                        .InitialValue = OriginColTag.EditRepNRHereInitialValue,
                                        .LastEditorValue = OriginColTag.EditRepNRHereInitialValue,
                                        .EditorFormat = OriginColTag.EditRepNRHereDataFormat,
                                        .LinkedComboBoxEdit = EditControl,
                                        .ParentBandedGridColumn = BGC
                                        }

                                    If Not DoneFirstItem Then

                                        MaxItemLength = GetMaxListItemLength(EditControl.RenderedListItems)
                                        CurrBaseCaption = New String("_", MaxItemLength + 2)
                                        DoneFirstItem = True

                                    End If



                                    RetComb.AccessibleName = "InheaderCombo_" & InHeaderCombosCount.ToString

                                    If OriginColTag.EditRepNRHereDataFormat = "S" Then

                                        RetComb.AppearanceDropDown.TextOptions.HAlignment = HorzAlignment.Center
                                        RetComb.Appearance.TextOptions.HAlignment = HorzAlignment.Center
                                        BGC.Width = BGC.Width * 2

                                    Else

                                        RetComb.AppearanceDropDown.TextOptions.HAlignment = HorzAlignment.Far
                                        RetComb.AppearanceDropDown.Font = EditControl.Font
                                        RetComb.Appearance.TextOptions.HAlignment = HorzAlignment.Far
                                        OriginColTag.ColWidthMultiplier = 1.33

                                    End If

                                    With BGC

                                        .AppearanceHeader.Options.UseBackColor = True
                                        'Keep the real header background consistent with ordinary headers.
                                        'ColumnInplaceEditorHelper paints AbovoComboBGC only inside the
                                        'single-line editor rectangle.
                                        .AppearanceHeader.BackColor = Color.White
                                        .AppearanceHeader.Options.UseForeColor = True
                                        .AppearanceHeader.ForeColor = AbovoBlue
                                        .Caption = CurrBaseCaption

                                    End With

                                    With RetComb

                                        .Appearance.Options.UseBorderColor = True
                                        .Appearance.BorderColor = AbovoComboBGC
                                        .Appearance.Options.UseTextOptions = True
                                        .Appearance.Options.UseBackColor = True
                                        .Appearance.BackColor = AbovoComboBGC
                                        .Appearance.ForeColor = Color.White

                                        .Appearance.Options.UseFont = True
                                        .Appearance.Font = New Font("Segoe UI", 12, FontStyle.Regular)
                                        '.Appearance.FontStyleDelta = FontStyle.Bold
                                        .Tag = InColumnEditorTag

                                    End With

                                    Dim DefAppearance As AppearanceObject = RetComb.Appearance

                                    RetComb.AppearanceDisabled.Assign(DefAppearance)
                                    RetComb.AppearanceFocused.Assign(DefAppearance)
                                    RetComb.AppearanceReadOnly.Assign(DefAppearance)

                                    'GridControls(GridCount).RepositoryItems.Add(RetComb)

                                    AddHandler RetComb.EditValueChanged, AddressOf ColumnHeaderEmbededComboChanged
                                    OriginColTag.InColumnEditorCombo = EditControl
                                    OriginColTag.HasIncolumnEditor = True
                                    Dim helper As ColumnInplaceEditorHelper = New ColumnInplaceEditorHelper(BGC, RetComb)
                                    InColumnEditorTag.InPlaceColumnHelper = helper
                                    EditControl.Tag = InColumnEditorTag
                                    helper.Tag = InColumnEditorTag
                                    helper.EditValue = OriginColTag.EditRepNRHereInitialValue
                                    helper.LinkedComboBoxEdit = EditControl
                                    EditControl.InPlaceColumnHelper = helper


                                ElseIf UCase(OriginColTag.EditRepNRHereEditor) = "DATE" Then

                                    InHeaderCombosCount += 1
                                    ReDim Preserve InHeaderCombos(InHeaderCombosCount)

                                    Dim EditControl As New AbovoDEHeaderDateBox
                                    EditControl.AddBlankFirstItem = True
                                    EditControl.InitialiseStandard(OriginColTag.EditRepNRHereComboRepository)
                                    EditControl.EditValue = OriginColTag.EditRepNRHereInitialValue

                                    Dim RetComb As RepositoryItemDateEdit = EditControl.Properties

                                    Dim InColumnEditorTag As New InColumnEditorTagDateEdit With {
                                        .EditingNRIndexPosition = OriginColTag.EditNRIndexPosition,
                                        .EditingNRName = OriginColTag.RepeatingNR,
                                        .NROrientation = IIf(OriginColTag.EditRepNRNROrientation = "PORT", Orientation.Vertical, Orientation.Horizontal),
                                        .EditorType = "DATE",
                                        .InitialValue = OriginColTag.EditRepNRHereInitialValue,
                                        .LastEditorValue = OriginColTag.EditRepNRHereInitialValue,
                                        .EditorFormat = OriginColTag.EditRepNRHereDataFormat,
                                        .LinkedDateBoxEdit = EditControl,
                                        .ParentBandedGridColumn = BGC
                                        }





                                    RetComb.AccessibleName = "InheaderCombo_" & InHeaderCombosCount.ToString

                                    If OriginColTag.EditRepNRHereDataFormat = "S" Then

                                        RetComb.AppearanceDropDown.TextOptions.HAlignment = HorzAlignment.Center
                                        RetComb.Appearance.TextOptions.HAlignment = HorzAlignment.Center
                                        BGC.Width = BGC.Width * 2

                                    Else

                                        RetComb.AppearanceDropDown.TextOptions.HAlignment = HorzAlignment.Far
                                        RetComb.AppearanceDropDown.Font = EditControl.Font
                                        RetComb.Appearance.TextOptions.HAlignment = HorzAlignment.Far
                                        OriginColTag.ColWidthMultiplier = 1.33

                                    End If

                                    With BGC

                                        .AppearanceHeader.Options.UseBackColor = True
                                        'Keep the real header background consistent with ordinary headers.
                                        'ColumnInplaceEditorHelper paints AbovoComboBGC only inside the
                                        'single-line editor rectangle.
                                        .AppearanceHeader.BackColor = Color.White
                                        .AppearanceHeader.Options.UseForeColor = True
                                        .AppearanceHeader.ForeColor = AbovoBlue
                                        .Caption = CurrBaseCaption

                                    End With

                                    With RetComb

                                        .Appearance.Options.UseBorderColor = True
                                        .Appearance.BorderColor = AbovoComboBGC
                                        .Appearance.Options.UseTextOptions = True
                                        .Appearance.Options.UseBackColor = True
                                        .Appearance.BackColor = AbovoComboBGC
                                        .Appearance.ForeColor = Color.White

                                        .Appearance.Options.UseFont = True
                                        .Appearance.Font = New Font("Segoe UI", 12, FontStyle.Regular)
                                        '.Appearance.FontStyleDelta = FontStyle.Bold
                                        .Tag = InColumnEditorTag

                                    End With

                                    Dim DefAppearance As AppearanceObject = RetComb.Appearance

                                    RetComb.AppearanceDisabled.Assign(DefAppearance)
                                    RetComb.AppearanceFocused.Assign(DefAppearance)
                                    RetComb.AppearanceReadOnly.Assign(DefAppearance)

                                    'GridControls(GridCount).RepositoryItems.Add(RetComb)

                                    AddHandler RetComb.EditValueChanged, AddressOf ColumnHeaderEmbededDateEChanged
                                    OriginColTag.InColumnEditorDate = EditControl
                                    OriginColTag.HasIncolumnEditor = True
                                    Dim helper As ColumnInplaceEditorHelper = New ColumnInplaceEditorHelper(BGC, RetComb)
                                    InColumnEditorTag.InPlaceColumnHelper = helper
                                    EditControl.Tag = InColumnEditorTag
                                    helper.Tag = InColumnEditorTag
                                    helper.EditValue = OriginColTag.EditRepNRHereInitialValue
                                    helper.LinkedDateEdit = EditControl
                                    EditControl.InPlaceColumnHelper = helper

                                End If
                            End If

                        End If

                    Next

                    BandExtendersCount += 1

                    'Dim Actor As New ActionToken With {.ActionType = "NREDIT", .ActionStrData1 = OriginColTag.ActionNR, .ActionNumber1 = SetSectionID, .ActionStrData2 = OriginColTag.BandEditDescription} ', .ActionData3 = ActiveDataSet.RepeatingHeaderText}
                    'If Not DontAddCDHDef Then AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomDrawColumnHeader, AddressOf DefaultCustomDrawColumnHeader
                    'AddHandler UsedGridVIEWS(GridViewCount).CustomDrawColumnHeader, AddressOf DefaultCustomDrawColumnHeader
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomDrawCell, AddressOf GridView_CustomDrawCell
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).ValidatingEditor, AddressOf GridView_ValidatingEditor
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).ShowingEditor, AddressOf GridView_ShowingEditor
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).ShownEditor, AddressOf GridView_ShownEditor
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).Click, AddressOf GridView_Event_SingleClick
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomRowCellEdit, AddressOf GridView_CustomCellEditor
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomRowCellEditForEditing, AddressOf GridView_RowCellEditForEditing
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CalcRowHeight, AddressOf GridView_CalcBandedGridRowHeight
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).ClipboardRowPasting, AddressOf GVPasting
                    UsedBANDedGridVIEWS(BandGridViewsCount).VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Auto

                    UsedBANDedGridVIEWS(BandGridViewsCount).EndUpdate()

                    UsedBANDedGridVIEWS(BandGridViewsCount).Name = "BandGridView_" & BandGridViewsCount.ToString
                    UsedBANDedGridVIEWS(BandGridViewsCount).AccessibleName = "BandGridView_" & BandGridViewsCount.ToString

                    GridControls(GridCount).ViewCollection.Add(UsedBANDedGridVIEWS(BandGridViewsCount))
                    GridControls(GridCount).MainView = UsedBANDedGridVIEWS(BandGridViewsCount)
                    UsedBANDedGridVIEWS(BandGridViewsCount).GridControl = GridControls(GridCount)



                    UsedBANDedGridVIEWS(BandGridViewsCount).OptionsView.BestFitMaxRowCount = ActiveDataSet.UsedRows
                    UnboundDataSources(UBSDataSourceCount).ActiveGridBandedView = UsedBANDedGridVIEWS(BandGridViewsCount)
                    UnboundDataSources(UBSDataSourceCount).InBandedMode = True

                Else

                    UnboundDataSources(UBSDataSourceCount).ActiveGridView = UsedGridVIEWS(GridViewCount)

                    'If Not DontAddCDHDef Then AddHandler UsedGridVIEWS(GridViewCount).CustomDrawColumnHeader, AddressOf DefaultCustomDrawColumnHeader

                    AddHandler UsedGridVIEWS(GridViewCount).CustomDrawCell, AddressOf GridView_CustomDrawCell
                    AddHandler UsedGridVIEWS(GridViewCount).ValidatingEditor, AddressOf GridView_ValidatingEditor
                    AddHandler UsedGridVIEWS(GridViewCount).ShowingEditor, AddressOf GridView_ShowingEditor
                    AddHandler UsedGridVIEWS(GridViewCount).ShownEditor, AddressOf GridView_ShownEditor
                    AddHandler UsedGridVIEWS(GridViewCount).CustomRowCellEdit, AddressOf GridView_CustomCellEditor
                    AddHandler UsedGridVIEWS(GridViewCount).Click, AddressOf GridView_Event_SingleClick
                    AddHandler UsedGridVIEWS(GridViewCount).CustomRowCellEditForEditing, AddressOf GridView_RowCellEditForEditing
                    ' AddHandler UsedGridVIEWS(GridViewCount).ColumnWidthChanged, AddressOf GridView_ColumnWidthChanged
                    AddHandler UsedGridVIEWS(GridViewCount).ClipboardRowPasting, AddressOf GVPasting
                    ExpanableGridControlsCount += 1
                    ReDim Preserve ExpanableGridControls(ExpanableGridControlsCount)
                    ExpanableGridControls(ExpanableGridControlsCount) = GridControls(GridCount)

                    UsedGridVIEWS(GridViewCount).VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Auto

                End If
#End Region



                If Not ApplyBands Then

                    BandColumButtonsAndDrawing = New CombinedColumnBandButton_HeaderFooterDrawer(UsedGridVIEWS(GridViewCount), Me, SetSectionID)

                End If



                If ActiveDataSet.RowExpandsByModel = "NRRI" OrElse
                   ActiveDataSet.RowExpandsByModel = "NRCI" Then

                    FooterOn = True

                    Dim FooterActionData As String =
                        ActiveDataSet.RowExpandByNR

                    Dim Actor As New ActionToken With {
                        .ActionType = ActiveDataSet.RowExpandsByModel,
                        .ActionStrData1 = FooterActionData,
                        .ActionNumber1 = SetSectionID,
                        .ActionDescription = "Edit " & FooterActionData
                    }

#If DEBUG Then
#End If

                    Dim itemCust As New GridColumnSummaryItem
                    itemCust.SummaryType = DevExpress.Data.SummaryItemType.Custom

                    If ApplyBands Then

                        itemCust.FieldName = UsedBANDedGridVIEWS(BandGridViewsCount).Columns(0).FieldName
                        UsedBANDedGridVIEWS(BandGridViewsCount).Columns(0).Summary.Add(itemCust)
                        BandColumButtonsAndDrawing.AddCustomAddRowsButton(UsedBANDedGridVIEWS(BandGridViewsCount).Columns(0), Actor)
                        Dim Col1Tag As DataColumnTag = UsedBANDedGridVIEWS(BandGridViewsCount).Columns(0).Tag
                        Col1Tag.HasIncolumnButton = True
                        If Not CustSummariesAddedToGrid Then
                            'AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomSummaryCalculate, AddressOf GridView_CustomSummaryCalculate
                            CustSummariesAddedToGrid = True
                        End If

                        If Not BandedFooterDone Then

                            BandedFooterDone = True
                            UsedBANDedGridVIEWS(BandGridViewsCount).OptionsView.ShowFooter = True
                            UsedBANDedGridVIEWS(BandGridViewsCount).FooterPanelHeight += (2 * DefaultGridCellPadding)

                        End If

                    Else


                        itemCust.FieldName = UsedGridVIEWS(GridViewCount).Columns(0).FieldName
                        UsedGridVIEWS(GridViewCount).Columns(0).Summary.Add(itemCust)
                        BandColumButtonsAndDrawing.AddCustomAddRowsButton(UsedGridVIEWS(GridViewCount).Columns(0), Actor)
                        Dim Col1Tag As DataColumnTag = UsedGridVIEWS(GridViewCount).Columns(0).Tag
                        Col1Tag.HasIncolumnButton = True
                        If Not CustSummariesAddedToGrid Then
                            'AddHandler UsedGridVIEWS(GridViewCount).CustomSummaryCalculate, AddressOf GridView_CustomSummaryCalculate
                            CustSummariesAddedToGrid = True
                        End If

                        If Not FooterDone Then
                            FooterOn = True
                            SetFooterOn(UsedGridVIEWS(GridViewCount), False)
                        End If

                    End If

                    DontAddCDHDef = True

                End If

                AddHandler GridControls(GridCount).DoubleClick, AddressOf VerifyDoubleClick

                Dim IdealGridHeight As Integer

                If ApplyBands Then

                    Dim GVI As GridViewInfo = UsedBANDedGridVIEWS(BandGridViewsCount).GetViewInfo()
                    Dim HeaderHeight As Integer = GVI.ColumnRowHeight
                    Dim FooterHeight As Integer = GVI.FooterCellHeight
                    IdealGridHeight = (IdealGridRowHeight + (2 * DefaultGridCellPadding)) * (ActiveDataSet.RowCount) + HeaderHeight + FooterHeight

                Else

                    Dim GVI As GridViewInfo = UsedGridVIEWS(GridViewCount).GetViewInfo()
                    Dim HeaderHeight As Integer = GVI.ColumnRowHeight
                    Dim FooterHeight As Integer = GVI.FooterCellHeight
                    IdealGridHeight = (IdealGridRowHeight + (2 * DefaultGridCellPadding)) * (ActiveDataSet.RowCount) + HeaderHeight + FooterHeight

                End If

                If IdealGridHeight > MaxGridHeight Then

                    IdealGridHeight = MaxGridHeight

                End If

                If SectionElement.GridControls IsNot Nothing Then

                    Dim LastRight As Integer = 0
                    Dim ButtText As String = ""
                    Dim NewButton As DevExpress.XtraEditors.SimpleButton = Nothing
                    Dim Layout As LayoutControl = New LayoutControl()
                    Layout.BackColor = Color.White

                    Layout.OptionsView.EnableTransparentBackColor = False
                    Layout.Padding = DefaultTablePanelPadding

                    Layout.Dock = DockStyle.Fill

                    Layout.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
                    Layout.LookAndFeel.UseDefaultLookAndFeel = False
                    Layout.OptionsView.EnableTransparentBackColor = True
                    Layout.Root.AppearanceGroup.BackColor = System.Drawing.Color.Transparent
                    Layout.Root.AppearanceGroup.Options.UseBackColor = True

                    TP.Controls.Add(Layout)
                    TP.SetCell(Layout, TPRowCount, 0)

                    Dim groupButtons As LayoutControlGroup = Layout.Root.AddGroup()
                    groupButtons.Name = "GroupButtons"
                    groupButtons.GroupBordersVisible = False
                    groupButtons.AppearanceGroup.BackColor = Color.White
                    groupButtons.AppearanceGroup.Options.UseBackColor = True

                    Dim lastItem As LayoutControlItem = Nothing
                    Dim DoneFirst As Boolean = False

                    For Each GridButtControl In SectionElement.GridControls

                        NewButton = New DevExpress.XtraEditors.SimpleButton()

                        Dim itemButton As LayoutControlItem = Nothing

                        If Not DoneFirst Then

                            itemButton = Layout.AddItem(groupButtons, InsertType.Left)

                            DoneFirst = True
                            lastItem = itemButton

                        Else

                            itemButton = Layout.AddItem(lastItem, InsertType.Right)
                            DoneFirst = True
                            lastItem = itemButton

                        End If

                        ButtText = GridButtControl.CommandText.Replace("vblf", "<br>")

                        NewButton.AllowHtmlDraw = DefaultBoolean.True
                        NewButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                        NewButton.Text = ButtText
                        NewButton.ToolTip = GridButtControl.CommandTip

                        GridButtControl.AttachedGrid = GridControls(GridCount)

                        NewButton.Tag = GridButtControl

                        Dim bestSize As Size = NewButton.CalcBestSize()
                        NewButton.Width = bestSize.Width
                        NewButton.Height = bestSize.Height

                        AddHandler NewButton.Click, AddressOf Grid_ButtonClick

                        itemButton.Control = NewButton

                    Next

                    Dim emptySpace As New EmptySpaceItem
                    emptySpace.Parent = groupButtons

                    SectionControlsCumlHeight += NewButton.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                    TPRowCount += 1

                    TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))

                    TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                End If

                GridControls(GridCount).Height = IdealGridHeight

                If IdealGridWidth > TP.Width - 50 Then IdealGridWidth = TP.Width - 50

                TP.Controls.Add(GridControls(GridCount))
                TP.SetCell(GridControls(GridCount), TPRowCount, 0)
                TP.SetColumnSpan(GridControls(GridCount), 4)

                SectionControlsCumlHeight += IdealGridHeight + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)
                TP.AutoSize = True
                TP.AutoSizeMode = AutoSizeMode.GrowAndShrink

                'GridControls(GridCount).Refresh()


#End Region

#Region "VGrid"

            ElseIf SectionElement.Type = "VGrid" Then

                'Joint Venture Assumptions is structurally different from a normal
                'DataCellRange VGrid.  The worksheet has Description/Year columns
                'followed by one physical column per JV.  Build that shape directly
                'from the named ranges so the VGrid mirrors Excel rather than showing
                'the helper/total ranges produced by MergeDownAndPivot.
                If CSID = 21 Then

                    ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)
                    RegisterDataSetDependencies(SetSectionID, ActiveDataSet)

                    Dim JVGrid As VGridControl =
                        BuildJointVentureExcelStyleVGrid(SetSectionID)

                    Dim JVIdealHeight As Integer =
                        Math.Max(300, GetPreferredVGridContentHeight(JVGrid) + 50)

                    JVGrid.Tag = New VGridLayoutTag With {
                        .TableRowIndex = TPRowCount
                    }

                    JVGrid.Height = JVIdealHeight

                    'Whole-grid structural command.  Adding a Joint Venture is
                    'a physical workbook COLUMN operation, so present it above
                    'the VGrid exactly like the Funding "Add Funding Columns"
                    'command rather than as an Add Rows button on the first category.
                    Dim JVButtonPanel As System.Windows.Forms.Panel =
                        CreateJointVentureAddColumnsPanel(
                            JVGrid,
                            SetSectionID)

                    TP.Controls.Add(JVButtonPanel)
                    TP.SetCell(JVButtonPanel, TPRowCount, 0)
                    TP.SetColumnSpan(JVButtonPanel, 4)

                    SectionControlsCumlHeight += JVButtonPanel.Height
                    TPRowCount += 1
                    TP.Rows.Add(
                        New TablePanelRow(
                            TablePanelEntityStyle.AutoSize,
                            50))
                    TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                    JVGrid.Tag = New VGridLayoutTag With {
                        .TableRowIndex = TPRowCount
                    }

                    TP.Controls.Add(JVGrid)
                    TP.SetCell(JVGrid, TPRowCount, 0)
                    TP.SetColumnSpan(JVGrid, 4)
                    TP.Rows(TPRowCount).Style = TablePanelEntityStyle.Absolute
                    TP.Rows(TPRowCount).Height = JVIdealHeight
                    SectionControlsCumlHeight += JVIdealHeight +
                        DefaultTablePanelPadding.Top +
                        DefaultTablePanelPadding.Bottom
                    TP.AutoSize = True
                    TP.AutoSizeMode = AutoSizeMode.GrowAndShrink

                    Continue For

                End If

                Dim ApplyBands As Boolean = DataPres.DataSets(SectionElement.ControlSourceIndex).HasBands
                Dim HasActions As Boolean = False
                Dim HasLockRules As Boolean = False
                Dim BandColumButtonsAndDrawing As CombinedColumnBandButton_HeaderFooterDrawer = Nothing
                Dim CustSummariesAddedToGrid As Boolean = False
                Dim BandedFooterDone As Boolean = False

                ColCount = -1

                FooterOn = False
                FooterDone = False

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)


                ColList = New List(Of String)

                PropertiesCount = -1 'reset
                UBSDataSourceCount += 1
                ReDim Preserve UnboundDataSources(UBSDataSourceCount)
                Dim SetTag As New AbovoUnboundSourceTag With {.GSID = GSID, .CSID = CSID, .RO = ActiveDataSet.RO, .DSIndex = SectionElement.ControlSourceIndex}

                UnboundDataSources(UBSDataSourceCount) = New AbovoUnboundSource(UBSDataSourceCount, SetTag)

                RegisterDataSetDependencies(SetSectionID, ActiveDataSet)


                Dim IdealGridWidth As Double = 0

                Dim ColTag As DataColumnTag = Nothing


                For Each PresentedColumn In ActiveDataSet.DataColumns

                    ColTag = PresentedColumn.ColumnTag

                    If ColTag.IsCalculated Then UnboundDataSources(UBSDataSourceCount).UBSTag.HasCalcs = True

                    ColCount += 1
                    ColName = "Col_" & ColCount.ToString
                    ColTag.ActiveColumnName = ColName
                    PropertiesCount += 1
                    ReDim Preserve PropertyArray(PropertiesCount)

                    Select Case ColTag.DataType

                        Case "S"
                            PropType = GetType(String)

                        Case "I", "Y", "D"
                            PropType = GetType(Integer)

                        Case "N", "P", "C", "M", "SM", "R"
                            PropType = GetType(Double)

                        Case "B"
                            PropType = GetType(Integer)

                        Case Else
                            PropType = GetType(String)

                    End Select


                    PropertyArray(PropertiesCount) = New UnboundSourceProperty With {
                        .UserTag = PresentedColumn.ColumnTag,
                        .DisplayName = " " & PresentedColumn.ColumnTag.ColumnHeading & " ",
                        .Name = ColName,
                        .PropertyType = PropType
                    }

                    ColList.Add(ColName)

                Next

                PropertyList = PropertyArray

                UnboundDataSources(UBSDataSourceCount).Properties.AddRange(PropertyList)
                UnboundDataSources(UBSDataSourceCount).InVertMode = True

                'Add dataaccess/push handlers
                AddHandler UnboundDataSources(UBSDataSourceCount).ValueNeeded, AddressOf UnboundDS_ValueNeeded
                AddHandler UnboundDataSources(UBSDataSourceCount).ValuePushed, AddressOf UnboundDS_ValuePushed

                'Create new grid for DS

                VertGridCount += 1
                ReDim Preserve VertGridControls(VertGridCount)



                VertGridControls(VertGridCount) = New VGridControl() With {
                    .Name = "VGridControl_" & VertGridCount.ToString,
                    .Parent = Me,
                    .Dock = DockStyle.None,
                    .Anchor = System.Windows.Forms.AnchorStyles.Top Or
                              System.Windows.Forms.AnchorStyles.Left
                }

                Dim VertGrid As VGridControl = VertGridControls(VertGridCount)

                CurrentAbovoTabPage.AddVGrid(VertGrid)
                AddHandler VertGrid.KeyDown, AddressOf VGrid_KeyDown
                AddHandler VertGrid.EditorKeyDown, AddressOf VGrid_KeyDown

                VertGrid.DataSource = UnboundDataSources(UBSDataSourceCount)

                Dim FillSize As Integer = ActiveDataSet.RowCount
                UnboundDataSources(UBSDataSourceCount).SetRowCount(FillSize)
                VertGrid.ForceInitialize()

                '=============================================================
                ' Joint Venture assumptions - Description + Year pairs
                '
                ' CSID 21 is the Joint Venture Assumptions child structure.
                ' Keep the worksheet / AbovoUnboundSource shape unchanged, but
                ' replace each eligible auto-generated Description + Year pair
                ' with one native DevExpress MultiEditorRow.
                '=============================================================
                Dim JVHandledVGridFields As New HashSet(Of String)(
                    StringComparer.Ordinal)

                If CSID = 21 Then

                    Dim JVTargetBands() As String = {
                        "Investments",
                        "Share of Surplus/Deficit",
                        "JV Cash Investment Repayments",
                        "JV Interest Rate"
                    }

                    For Each JVBandName As String In JVTargetBands

                        Dim JVDescriptionTag As DataColumnTag = Nothing
                        Dim JVYearTag As DataColumnTag = Nothing

                        For Each JVPresentedColumn In ActiveDataSet.DataColumns

                            Dim JVCandidateTag As DataColumnTag =
                                JVPresentedColumn.ColumnTag

                            If JVCandidateTag Is Nothing Then Continue For

                            If Not String.Equals(
                                JVCandidateTag.BandID,
                                JVBandName,
                                StringComparison.Ordinal) Then

                                Continue For

                            End If

                            If String.Equals(
                                JVCandidateTag.ColumnHeading.Trim(),
                                "Description",
                                StringComparison.OrdinalIgnoreCase) Then

                                JVDescriptionTag = JVCandidateTag

                            ElseIf String.Equals(
                                JVCandidateTag.ColumnHeading.Trim(),
                                "Year",
                                StringComparison.OrdinalIgnoreCase) Then

                                JVYearTag = JVCandidateTag

                            End If

                        Next

                        'Some JV bands may be supplied by another interface
                        'datasource.  Transform only a complete pair present in
                        'this VGrid's ActiveDataSet.
                        If JVDescriptionTag Is Nothing OrElse
                           JVYearTag Is Nothing Then

                            Continue For

                        End If

                        Dim JVDescriptionRow As BaseRow =
                            VertGrid.Rows.GetRowByFieldName(
                                JVDescriptionTag.ActiveColumnName,
                                True)

                        Dim JVYearRow As BaseRow =
                            VertGrid.Rows.GetRowByFieldName(
                                JVYearTag.ActiveColumnName,
                                True)

                        If JVDescriptionRow Is Nothing OrElse
                           JVYearRow Is Nothing Then

                            Continue For

                        End If

                        'The current XML does not need to be changed for this
                        'stage.  Supply the standard OrdinalYears repository to
                        'the Year item at runtime.
                        If String.IsNullOrWhiteSpace(
                            JVYearTag.RepositaryID) Then

                            JVYearTag.RepositaryID =
                                "Rep_OrdinalYears"

                        End If

                        JVYearTag.HasComboEdit = True

                        Dim JVDescriptionEdit As New RepositoryItemTextEdit()

                        JVDescriptionEdit.Name =
                            "JVDescription_" &
                            VertGridCount.ToString &
                            "_" &
                            JVBandName.Replace(" ", "_").
                                       Replace("/", "_")

                        VertGrid.RepositoryItems.Add(
                            JVDescriptionEdit)

                        JVDescriptionTag.DefaultTextEditor =
                            JVDescriptionEdit

                        Dim JVYearEditControl As RepositaryItems.AbovoRespositaryItem =
                            RepositaryItems.GetEditor(
                                JVYearTag.RepositaryID,
                                ModelID)

                        If JVYearEditControl Is Nothing OrElse
                           JVYearEditControl.RepType <> "CMB" OrElse
                           JVYearEditControl.RetCombo Is Nothing Then


                            Continue For

                        End If

                        GridCombosCount += 1

                        JVYearEditControl.RetCombo.Name =
                            "JVYearCombo_" &
                            GridCombosCount.ToString

                        AddHandler JVYearEditControl.RetCombo.Enter,
                            AddressOf ComboOpen

                        VertGrid.RepositoryItems.Add(
                            JVYearEditControl.RetCombo)

                        ComboReposClassesCount += 1
                        ReDim Preserve ComboReposClasses(ComboReposClassesCount)

                        ComboReposClasses(ComboReposClassesCount) =
                            New AbovoGridRespoitaryCombo With {
                                .ID = ComboReposClassesCount,
                                .RepoistaryID = JVYearTag.RepositaryID,
                                .Combo = JVYearEditControl.RetCombo,
                                .ModelID = ModelID,
                                .VGridID = VertGridCount
                            }

                        CurrentAbovoTabPage.AddRepCombo(
                            ComboReposClasses(ComboReposClassesCount))

                        Dim JVDescriptionProperties As New MultiEditorRowProperties()

                        With JVDescriptionProperties
                            .Caption = JVBandName.Replace("vblf", vbLf)
                            .FieldName = JVDescriptionTag.ActiveColumnName
                            .ReadOnly =
                                JVDescriptionTag.IsReadOnly OrElse
                                JVDescriptionTag.IsCalculated
                            .RowEdit = JVDescriptionEdit
                        End With

                        Dim JVYearProperties As New MultiEditorRowProperties()

                        With JVYearProperties
                            .Caption = "Year"
                            .FieldName = JVYearTag.ActiveColumnName
                            .ReadOnly =
                                JVYearTag.IsReadOnly OrElse
                                JVYearTag.IsCalculated
                            .RowEdit = JVYearEditControl.RetCombo
                        End With

                        Dim JVMultiRow As New MultiEditorRow()

                        With JVMultiRow
                            .Name =
                                "JVPair_" &
                                JVBandName.Replace(" ", "_").
                                           Replace("/", "_")
                            .Height = IdealGridRowHeight
                            .Tag =
                                New VGridMultiEditorRowTag(
                                    JVDescriptionTag,
                                    JVYearTag)
                        End With

                        JVMultiRow.PropertiesCollection.AddRange(
                            New MultiEditorRowProperties() {
                                JVDescriptionProperties,
                                JVYearProperties
                            })

                        'Remove the two rows generated by ForceInitialize and
                        'insert their combined replacement at the first row's
                        'position.
                        Dim JVInsertIndex As Integer =
                            Math.Min(
                                VertGrid.Rows.IndexOf(JVDescriptionRow),
                                VertGrid.Rows.IndexOf(JVYearRow))

                        VertGrid.Rows.Remove(JVDescriptionRow)
                        VertGrid.Rows.Remove(JVYearRow)

                        If JVInsertIndex < 0 OrElse
                           JVInsertIndex > VertGrid.Rows.Count Then

                            VertGrid.Rows.Add(JVMultiRow)

                        Else

                            VertGrid.Rows.Insert(
                                JVMultiRow,
                                JVInsertIndex)

                        End If

                        JVHandledVGridFields.Add(
                            JVDescriptionTag.ActiveColumnName)

                        JVHandledVGridFields.Add(
                            JVYearTag.ActiveColumnName)

                    Next

                End If


                Dim LastBandID As String = Nothing
                Dim CatRow As CategoryRow = Nothing
                Dim IndexPos As Integer = 0
                Dim LastColour As Color = Color.Red
                Dim VGridCategories As New List(Of CategoryRow)
                Dim VGridInEditorAddingMode As Boolean = False

                'The rows have already been generated by ForceInitialize().
                'Find each row by its bound field name rather than relying upon Rows(0).
                For Each PresentedColumn In ActiveDataSet.DataColumns

                    ColTag = PresentedColumn.ColumnTag

                    'Description/Year fields already incorporated into a Joint
                    'Venture MultiEditorRow must not be recreated as ordinary
                    'EditorRows or CategoryRows.
                    If JVHandledVGridFields.Contains(
                        ColTag.ActiveColumnName) Then

                        IndexPos += 1
                        Continue For

                    End If

                    '-------------------------------------------------------------
                    ' Create category/band
                    '-------------------------------------------------------------
                    If IndexPos = 0 OrElse Not String.Equals(ColTag.BandID, LastBandID, StringComparison.Ordinal) Then

                        LastBandID = ColTag.BandID

                        'Mirror the XtraGrid in-header editor behaviour at band level.
                        'If the first field in a band carries EditRepNRHere, each
                        'repeating field/row in that band gets an editor in its row
                        'header (the first visual VGrid column).
                        VGridInEditorAddingMode = ColTag.EditRepNRHere

                        'Do not leave an unbanded row attached to the previous category.
                        If String.IsNullOrEmpty(ColTag.BandID) Then

                            CatRow = Nothing

                        Else

                            If LastColour = Color.Red Then

                                LastColour = Color.White

                            ElseIf LastColour = Color.White Then

                                LastColour = Color.Wheat

                            ElseIf LastColour = Color.Wheat Then

                                LastColour = AbovoBlue

                            ElseIf LastColour = AbovoBlue Then

                                LastColour = Color.Wheat

                            End If

                            CatRow = New CategoryRow("Category_" & IndexPos.ToString) With {
                                .Height = 40
                            }

                            Dim CategoryCaption As String = ColTag.BandID.Replace("vblf", vbLf)

                            If ColTag.BandID = "FixedLeft" Then CategoryCaption = " "

                            CatRow.Properties.Caption = CategoryCaption

                            Dim NewCategoryTag As New BandTag With {
                                .ID = VertGridCount,
                                .ActionNR = ColTag.ActionNR,
                                .ActionDescription = ColTag.BandEditDescription,
                                .HighLightColour = LastColour,
                                .DoBorder = True,
                                .ButtonObjectState = ObjectState.Normal
                            }

                            CatRow.Tag = NewCategoryTag
                            VertGrid.Rows.Add(CatRow)
                            VGridCategories.Add(CatRow)

                        End If

                    End If

                    '-------------------------------------------------------------
                    ' Find the auto-created row belonging to this data property
                    '-------------------------------------------------------------
                    Dim VRow As EditorRow =
        TryCast(
            VertGrid.Rows.GetRowByFieldName(
                ColTag.ActiveColumnName,
                True),
            EditorRow)

                    If VRow Is Nothing Then

                        'Shouldn't normally happen after ForceInitialize(), but this makes
                        'the code safe if auto row generation changes.
                        VRow = New EditorRow(ColTag.ActiveColumnName)

                        VertGrid.Rows.Add(VRow)

                    End If

                    With VRow

                        .Tag = ColTag
                        .Height = IdealGridRowHeight

                        With .Properties

                            .FieldName = ColTag.ActiveColumnName
                            .Caption = " " & ColTag.ColumnHeading & " "
                            .ReadOnly = ColTag.IsReadOnly OrElse ColTag.IsCalculated

                        End With

                    End With

                    '-------------------------------------------------------------
                    ' Category/band action - VGrid equivalent of the XtraGrid
                    ' GridBand action button.  Use the first actionable member of
                    ' the category as the source of the repeating-NR action.
                    '-------------------------------------------------------------
                    If CatRow IsNot Nothing AndAlso ColTag.HasActions Then

                        Dim CategoryTag As BandTag = TryCast(CatRow.Tag, BandTag)

                        If CategoryTag IsNot Nothing AndAlso Not CategoryTag.HasActions Then

                            Dim EditText As String

                            If ColTag.BandEditDescription IsNot Nothing Then

                                EditText = ColTag.BandEditDescription

                            Else

                                EditText = "Edit " & ColTag.ActiveColumnName & " of " & ColTag.BandID

                            End If

                            Dim ActToken As New ActionToken With {
                                .ActionType = ColTag.EditRepNRHereExpansionMethod,
                                .ActionStrData1 = ColTag.ActionNR,
                                .ActionNR = ColTag.ActionNR,
                                .ActionNumber1 = SetSectionID,
                                .ActionStrData2 = EditText,
                                .ActionDescription = EditText
                            }

                            CategoryTag.HasActions = True
                            CategoryTag.ActionToken = ActToken
                            CategoryTag.ActionNR = ColTag.ActionNR
                            CategoryTag.ActionDescription = EditText

                            VertGrid.InvalidateRow(CatRow)

                        End If

                    End If

                    '-------------------------------------------------------------
                    ' Appearance - equivalent to Grid column configuration
                    '-------------------------------------------------------------
                    With VRow.AppearanceCell

                        .Options.UseBackColor = True
                        .Options.UseForeColor = True
                        .Options.UseBorderColor = True

                        .BackColor = Color.Silver
                        .ForeColor = Color.Black
                        .BorderColor = Color.Silver

                    End With

                    If ColTag.IsReadOnly OrElse ColTag.IsCalculated Then

                        With VRow.AppearanceCell

                            .BackColor = Color.WhiteSmoke
                            .ForeColor = AbovoBlue
                            .BorderColor = Color.White

                        End With

                    End If

                    '-------------------------------------------------------------
                    ' Tool tip
                    '-------------------------------------------------------------
                    If Not IsNothing(ColTag.TipText) Then

                        VRow.Properties.Caption = " " & ColTag.ColumnHeading & " "

                        'Depending upon your DevExpress version BaseRow may expose
                        'ToolTip directly. If yours does, uncomment:
                        '
                        'VRow.ToolTip = TrimSpaces(ColTag.TipText)

                    End If

                    '-------------------------------------------------------------
                    ' Repository/combo editor
                    ' Same logic as GridColumn.ColumnEdit
                    '-------------------------------------------------------------
                    Dim UseCombo As Boolean = False

                    If Not IsNothing(ColTag.RepositaryID) Then

                        ColTag.HasComboEdit = True
                        UseCombo = True

                        Dim EditControl As RepositaryItems.AbovoRespositaryItem =
            RepositaryItems.GetEditor(
                ColTag.RepositaryID,
                ModelID)

                        If EditControl.RepType = "CMB" Then

                            GridCombosCount += 1

                            EditControl.RetCombo.Name =
                "VGridCombo_" & GridCombosCount.ToString

                            AddHandler EditControl.RetCombo.Enter, AddressOf ComboOpen
                            VertGrid.RepositoryItems.Add(EditControl.RetCombo)

                            ComboReposClassesCount += 1
                            ReDim Preserve ComboReposClasses(ComboReposClassesCount)

                            ComboReposClasses(ComboReposClassesCount) = New AbovoGridRespoitaryCombo With {
                                .ID = ComboReposClassesCount,
                                .RepoistaryID = ColTag.RepositaryID,
                                .Combo = EditControl.RetCombo,
                                .ModelID = ModelID,
                                .VGridID = VertGridCount}

                            CurrentAbovoTabPage.AddRepCombo(ComboReposClasses(ComboReposClassesCount))

                            'VGrid equivalent of GVcolumn.ColumnEdit
                            VRow.Properties.RowEdit = EditControl.RetCombo

                        End If

                    End If

                    '-------------------------------------------------------------
                    ' Data-type formatting/editors
                    ' Mirrors the Grid Select Case
                    '-------------------------------------------------------------
                    Select Case ColTag.DataType

                        Case "S"

                            VRow.AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Near

                            VRow.Properties.DisplayFormat.FormatType =
                DevExpress.Utils.FormatType.None

                            ColTag.ShowDefaultmask = -1


                        Case "M"

                            VRow.AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.DisplayFormat.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.DisplayFormat.FormatString = "c0"

                            Dim edit As New RepositoryItemTextEdit()

                            edit.Mask.MaskType =
                DevExpress.XtraEditors.Mask.MaskType.Numeric

                            edit.Mask.EditMask = "c5"

                            VertGrid.RepositoryItems.Add(edit)
                            VRow.Properties.RowEdit = edit

                            ColTag.DefaultTextEditor = edit
                            ColTag.ShowDefaultmask = 0


                        Case "SM"

                            VRow.AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.DisplayFormat.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.DisplayFormat.FormatString = "c2"

                            Dim edit As New RepositoryItemTextEdit()

                            edit.UseMaskAsDisplayFormat = True

                            edit.Mask.MaskType =
                DevExpress.XtraEditors.Mask.MaskType.Numeric

                            edit.Mask.EditMask = "c2"

                            VertGrid.RepositoryItems.Add(edit)
                            VRow.Properties.RowEdit = edit

                            ColTag.DefaultTextEditor = edit
                            ColTag.ShowDefaultmask = 1


                        Case "I"

                            VRow.AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.DisplayFormat.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.DisplayFormat.FormatString = "#,##0"

                            ColTag.ShowDefaultmask = -1

                            If Not UseCombo Then

                                Dim edit As New RepositoryItemTextEdit()

                                edit.Mask.MaskType =
                    DevExpress.XtraEditors.Mask.MaskType.Numeric

                                edit.Mask.EditMask = "n0"

                                VertGrid.RepositoryItems.Add(edit)
                                VRow.Properties.RowEdit = edit

                                ColTag.DefaultTextEditor = edit

                            End If


                        Case "C"

                            VRow.AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.DisplayFormat.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.DisplayFormat.FormatString = "#,###,##0"


                        Case "R"

                            VRow.AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.DisplayFormat.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.DisplayFormat.FormatString = "#,##0.00"

                            Dim edit As New RepositoryItemTextEdit()

                            edit.Mask.MaskType =
                DevExpress.XtraEditors.Mask.MaskType.Numeric

                            edit.Mask.EditMask = "n5"

                            VertGrid.RepositoryItems.Add(edit)
                            VRow.Properties.RowEdit = edit

                            ColTag.DefaultTextEditor = edit


                        Case "P"

                            VRow.AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.DisplayFormat.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.DisplayFormat.FormatString = "p2"

                            ColTag.ShowDefaultmask = 3

                            Dim SetMinVal As Object =
                If(ColTag.MinVal Is Nothing,
                   0,
                   If(ColTag.MinVal = "NOMIN",
                      Nothing,
                      CDec(ColTag.MinVal)))

                            Dim SetMaxVal As Object =
                If(ColTag.MaxVal Is Nothing,
                   1,
                   If(ColTag.MaxVal = "NOMAX",
                      Nothing,
                      CDec(ColTag.MaxVal)))

                            If Not ColTag.IsReadOnly Then

                                Dim edit As New RepositoryItemSpinEdit()

                                If SetMinVal IsNot Nothing Then
                                    edit.MinValue = CDec(SetMinVal)
                                End If

                                edit.Increment =
                    If(ColTag.DefIncrement Is Nothing,
                       CDec(0.0025),
                       CDec(ColTag.DefIncrement))

                                If SetMaxVal IsNot Nothing Then
                                    edit.MaxValue = CDec(SetMaxVal)
                                End If

                                edit.Mask.MaskType =
                    DevExpress.XtraEditors.Mask.MaskType.Numeric

                                edit.Mask.EditMask = "p2"

                                edit.DisplayFormat.FormatType =
                    DevExpress.Utils.FormatType.Numeric

                                edit.DisplayFormat.FormatString = "p2"

                                VertGrid.RepositoryItems.Add(edit)
                                VRow.Properties.RowEdit = edit

                                ColTag.DefaultTextEditor = edit

                            End If


                        Case "B"

                            'At present your XtraGrid block doesn't explicitly
                            'assign a Boolean editor either, so leave the VGrid
                            'using the DevExpress default.


                        Case Else

                            'Default editor generated from the underlying type.

                    End Select

                    '-------------------------------------------------------------
                    ' Repeating-NR editor in the VGrid row header
                    '-------------------------------------------------------------
                    If VGridInEditorAddingMode AndAlso
                       Not ColTag.EditRepNRHereROColumn AndAlso
                       Not String.IsNullOrEmpty(ColTag.BandID) Then

                        If UCase(ColTag.EditRepNRHereEditor) = "COMBO" Then

                            Dim EditControl As New AbovoDEHeaderComboBox
                            EditControl.AddBlankFirstItem = True
                            EditControl.InitialiseStandard(ColTag.EditRepNRHereComboRepository)
                            EditControl.EditValue = ColTag.EditRepNRHereInitialValue

                            Dim RetComb As RepositoryItemComboBox = EditControl.Properties

                            Dim InColumnEditorTag As New InColumnEditorTagCombo With {
                                .EditingNRIndexPosition = ColTag.EditNRIndexPosition,
                                .EditingNRName = ColTag.RepeatingNR,
                                .NROrientation = IIf(ColTag.EditRepNRNROrientation = "PORT", Orientation.Vertical, Orientation.Horizontal),
                                .EditorType = "COMBO",
                                .InitialValue = ColTag.EditRepNRHereInitialValue,
                                .LastEditorValue = ColTag.EditRepNRHereInitialValue,
                                .EditorFormat = ColTag.EditRepNRHereDataFormat,
                                .LinkedComboBoxEdit = EditControl
                            }

                            With RetComb
                                .Appearance.Options.UseBorderColor = True
                                .Appearance.BorderColor = AbovoComboBGC
                                .Appearance.Options.UseTextOptions = True
                                .Appearance.Options.UseBackColor = True
                                .Appearance.BackColor = AbovoComboBGC
                                .Appearance.ForeColor = Color.White
                                .Appearance.Options.UseFont = True
                                .Appearance.Font = New Font("Segoe UI", 12, FontStyle.Regular)
                                .Tag = InColumnEditorTag
                            End With

                            Dim DefAppearance As AppearanceObject = RetComb.Appearance
                            RetComb.AppearanceDisabled.Assign(DefAppearance)
                            RetComb.AppearanceFocused.Assign(DefAppearance)
                            RetComb.AppearanceReadOnly.Assign(DefAppearance)

                            'For a VGrid the RepositoryItem is a paint/template object only.
                            'Wire changes to the temporary live BaseEdit in the helper instead.
                            Dim Helper As New VGridRowInplaceEditorHelper(VertGrid,
                                                                         VRow,
                                                                         RetComb,
                                                                         AddressOf ColumnHeaderEmbededComboChanged)
                            Helper.Tag = InColumnEditorTag
                            Helper.EditValue = ColTag.EditRepNRHereInitialValue
                            Helper.LinkedComboBoxEdit = EditControl

                            InColumnEditorTag.InPlaceVGridRowHelper = Helper
                            EditControl.Tag = InColumnEditorTag
                            ColTag.InColumnEditorCombo = EditControl
                            ColTag.HasIncolumnEditor = True

                            Dim SectionHelpers As List(Of VGridRowInplaceEditorHelper) = Nothing
                            If Not VGridInplaceEditorHelpersBySection.TryGetValue(SetSectionID, SectionHelpers) Then
                                SectionHelpers = New List(Of VGridRowInplaceEditorHelper)
                                VGridInplaceEditorHelpersBySection(SetSectionID) = SectionHelpers
                            End If
                            SectionHelpers.Add(Helper)

                        ElseIf UCase(ColTag.EditRepNRHereEditor) = "DATE" Then

                            Dim EditControl As New AbovoDEHeaderDateBox
                            EditControl.AddBlankFirstItem = True
                            EditControl.InitialiseStandard(ColTag.EditRepNRHereComboRepository)
                            EditControl.EditValue = ColTag.EditRepNRHereInitialValue

                            Dim RetDate As RepositoryItemDateEdit = EditControl.Properties

                            Dim InColumnEditorTag As New InColumnEditorTagDateEdit With {
                                .EditingNRIndexPosition = ColTag.EditNRIndexPosition,
                                .EditingNRName = ColTag.RepeatingNR,
                                .NROrientation = IIf(ColTag.EditRepNRNROrientation = "PORT", Orientation.Vertical, Orientation.Horizontal),
                                .EditorType = "DATE",
                                .InitialValue = ColTag.EditRepNRHereInitialValue,
                                .LastEditorValue = ColTag.EditRepNRHereInitialValue,
                                .EditorFormat = ColTag.EditRepNRHereDataFormat,
                                .LinkedDateBoxEdit = EditControl
                            }

                            With RetDate
                                .Appearance.Options.UseBorderColor = True
                                .Appearance.BorderColor = AbovoComboBGC
                                .Appearance.Options.UseTextOptions = True
                                .Appearance.Options.UseBackColor = True
                                .Appearance.BackColor = AbovoComboBGC
                                .Appearance.ForeColor = Color.White
                                .Appearance.Options.UseFont = True
                                .Appearance.Font = New Font("Segoe UI", 12, FontStyle.Regular)
                                .Tag = InColumnEditorTag
                            End With

                            Dim DefAppearance As AppearanceObject = RetDate.Appearance
                            RetDate.AppearanceDisabled.Assign(DefAppearance)
                            RetDate.AppearanceFocused.Assign(DefAppearance)
                            RetDate.AppearanceReadOnly.Assign(DefAppearance)

                            'For a VGrid the RepositoryItem is a paint/template object only.
                            'Wire changes to the temporary live BaseEdit in the helper instead.
                            Dim Helper As New VGridRowInplaceEditorHelper(VertGrid,
                                                                         VRow,
                                                                         RetDate,
                                                                         AddressOf ColumnHeaderEmbededDateEChanged)
                            Helper.Tag = InColumnEditorTag
                            Helper.EditValue = ColTag.EditRepNRHereInitialValue
                            Helper.LinkedDateEdit = EditControl

                            InColumnEditorTag.InPlaceVGridRowHelper = Helper
                            EditControl.Tag = InColumnEditorTag
                            ColTag.InColumnEditorDate = EditControl
                            ColTag.HasIncolumnEditor = True

                            Dim SectionHelpers As List(Of VGridRowInplaceEditorHelper) = Nothing
                            If Not VGridInplaceEditorHelpersBySection.TryGetValue(SetSectionID, SectionHelpers) Then
                                SectionHelpers = New List(Of VGridRowInplaceEditorHelper)
                                VGridInplaceEditorHelpersBySection(SetSectionID) = SectionHelpers
                            End If
                            SectionHelpers.Add(Helper)

                        End If

                    End If

                    '-------------------------------------------------------------
                    ' Move the configured editor row beneath its category
                    '-------------------------------------------------------------
                    If CatRow IsNot Nothing Then

                        VertGrid.MoveRow(VRow, CatRow, False)

                    End If

                    IndexPos += 1

                Next

                '-------------------------------------------------------------
                ' VGrid category footers
                '
                'VGrid has no native per-category footer.  Add the synthetic
                'footer only after all real rows have been moved beneath their
                'categories, ensuring the footer remains the final child row.
                'The extender paints the category highlight and footer, and
                'draws the Add Rows action only where BandTag.HasActions is True.
                '-------------------------------------------------------------
                For Each Category As CategoryRow In VGridCategories

                    Dim CategoryTag As BandTag = TryCast(Category.Tag, BandTag)

                    Dim CategoryExtender As New VGridCategoryButtonExtender(
                        VertGrid,
                        Category,
                        Me,
                        "",
                        SetSectionID,
                        Category.Properties.Caption,
                        If(CategoryTag Is Nothing, Nothing, CategoryTag.ActionToken))

                    CategoryExtender.AddCustomButton()
                    VGridCategoryExtenders.Add(CategoryExtender)

                    Dim SectionExtenders As List(Of VGridCategoryButtonExtender) = Nothing
                    If Not VGridCategoryExtendersBySection.TryGetValue(SetSectionID, SectionExtenders) Then
                        SectionExtenders = New List(Of VGridCategoryButtonExtender)
                        VGridCategoryExtendersBySection(SetSectionID) = SectionExtenders
                    End If
                    SectionExtenders.Add(CategoryExtender)

                Next

                If ActiveDataSet.HasValidations Then


                    For Each ValList In ActiveDataSet.ValidationLists

                        VertGrid.RepositoryItems.Add(
            RepositaryItems.GetEditorFromList(ValList).RetCombo)

                    Next

                End If

                If ActiveDataSet.RO Then VertGrid.OptionsBehavior.Editable = False

                VertGrid.RepositoryItems.Add(CustCalcEditInteger)
                VertGrid.RepositoryItems.Add(CustCalcEditDecimal)

                Dim RowIndex As Integer = 0

                'For Each ColCheck As UnboundSourceProperty In PropertyArray

                '    'With VertGrid.Rows(RowIndex)

                '    '    .Tag = ColCheck
                '    '    .AppearanceCell.Options.UseBackColor = True
                '    '    .AppearanceCell.Options.UseBorderColor = True
                '    '    .AppearanceCell.BorderColor = System.Drawing.Color.Silver
                '    '    .AppearanceCell.BackColor = System.Drawing.Color.Silver
                '    '    .AppearanceCell.ForeColor = System.Drawing.Color.Black

                '    'End With



                '    'If ColCheck.UserTag.IsReadOnly OrElse ColCheck.UserTag.IsCalculated Then

                '    '    With VertGrid.Rows(RowIndex)

                '    '        '.OptionsRow.TabStop = False
                '    '        '.OptionsColumn.ReadOnly = True
                '    '        .AppearanceCell.BorderColor = System.Drawing.Color.White
                '    '        .AppearanceCell.BackColor = System.Drawing.Color.White
                '    '        .AppearanceCell.ForeColor = System.Drawing.Color.DarkGray

                '    '    End With

                '    'End If

                '    'If ColCheck.UserTag.IsFixed Then

                '    '    UsedGridVIEWS(GridViewCount).OptionsView.ColumnAutoWidth = False
                '    '    UsedGridVIEWS(GridViewCount).Columns(ColCheck.UserTag.ActiveColumnName).Fixed = FixedStyle.Left

                '    'End If
                '    RowIndex += 1

                'Next

                Formatter.FormatVertGrid(VertGrid)

                '--------------------------------------------------------------
                'Optional VGrid command strip.
                '
                'Most structural grids now use their footer action. Funding is
                'deliberately different: workbook facilities are physical
                'columns, while MergeDownAndPivot + VGrid presents them through
                'two pivots.  A conventional "Add Funding Columns" command is
                'therefore clearer than pretending the VGrid has addable rows.
                '
                'The command strip is generic so the same pattern can be reused
                'for the other exceptional VGrid interface later.
                '--------------------------------------------------------------
                If SectionElement.GridControls IsNot Nothing AndAlso
                   SectionElement.GridControls.Count > 0 Then

                    Dim VGridButtonPanel As New System.Windows.Forms.Panel With {
                        .AutoSize = False,
                        .Dock = System.Windows.Forms.DockStyle.None,
                        .Anchor = System.Windows.Forms.AnchorStyles.Top Or
                                  System.Windows.Forms.AnchorStyles.Left Or
                                  System.Windows.Forms.AnchorStyles.Right,
                        .Tag = VertGrid,
                        .Margin = New System.Windows.Forms.Padding(0),
                        .Padding = New System.Windows.Forms.Padding(
                            CInt(DefaultTablePanelPadding.Left),
                            CInt(DefaultTablePanelPadding.Top),
                            CInt(DefaultTablePanelPadding.Right),
                            CInt(DefaultTablePanelPadding.Bottom))
                    }

                    Dim VGridButtonHeight As Integer = 0

                    For Each VGridButtControl In SectionElement.GridControls

                        Dim VGridCommandButton As New DevExpress.XtraEditors.SimpleButton()

                        Dim VGridButtonText As String =
                            VGridButtControl.CommandText.Replace("vblf", "<br>")

                        VGridCommandButton.AllowHtmlDraw = DefaultBoolean.True
                        VGridCommandButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                        VGridCommandButton.Text = VGridButtonText
                        VGridCommandButton.ToolTip = VGridButtControl.CommandTip
                        VGridCommandButton.Tag = VGridButtControl

                        'Keep an explicit vertical-grid owner for this special
                        'command family. Add-only funding currently does not need
                        'the owner, but future VGrid commands may.
                        VGridButtControl.AttachedVGrid = VertGrid

                        Dim VGridButtonBestSize As Size =
                            VGridCommandButton.CalcBestSize()

                        VGridCommandButton.Width =
                            Math.Max(VGridButtonBestSize.Width, 140)

                        VGridCommandButton.Height =
                            VGridButtonBestSize.Height

                        If VGridCommandButton.Height > VGridButtonHeight Then
                            VGridButtonHeight = VGridCommandButton.Height
                        End If

                        AddHandler VGridCommandButton.Click,
                            AddressOf Grid_ButtonClick

                        VGridButtonPanel.Controls.Add(VGridCommandButton)

                    Next

                    VGridButtonPanel.Height =
                        VGridButtonHeight +
                        DefaultTablePanelPadding.Top +
                        DefaultTablePanelPadding.Bottom

                    TP.Controls.Add(VGridButtonPanel)
                    TP.SetCell(VGridButtonPanel, TPRowCount, 0)
                    TP.SetColumnSpan(VGridButtonPanel, 4)

                    SectionControlsCumlHeight +=
                        VGridButtonHeight +
                        DefaultTablePanelPadding.Top +
                        DefaultTablePanelPadding.Bottom

                    TPRowCount += 1
                    TP.Rows.Add(
                        New TablePanelRow(
                            TablePanelEntityStyle.AutoSize,
                            50))

                    TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                End If

                Dim IdealGridHeight As Integer = (IdealGridRowHeight + (2 * DefaultGridCellPadding)) * (FillSize) + 50

                'Remember which TablePanel row owns this VGrid.  The final layout
                'pass will replace this provisional height with the actual remaining
                'tab-page viewport height.
                VertGrid.Tag = New VGridLayoutTag With {
                    .TableRowIndex = TPRowCount
                }

                VertGrid.Height = IdealGridHeight
                TP.Controls.Add(VertGrid)
                TP.SetCell(VertGrid, TPRowCount, 0)
                TP.SetColumnSpan(VertGrid, 4)

                'This row must NOT be AutoSize. DevExpress documents that an
                'AutoSize TablePanelRow ignores its Height property and sizes from
                'content. That was the real reason the VGrid stopped at ~60-70%
                'even after we assigned VG.Height to the full viewport.
                TP.Rows(TPRowCount).Style = TablePanelEntityStyle.Absolute
                TP.Rows(TPRowCount).Height = IdealGridHeight

                SectionControlsCumlHeight += IdealGridHeight + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)
                TP.AutoSize = True
                TP.AutoSizeMode = AutoSizeMode.GrowAndShrink
                VertGrid.BestFit()
                UnboundDataSources(UBSDataSourceCount).AttachedVertGrid = VertGrid

                AddHandler VertGrid.CustomDrawRowValueCell, AddressOf VGrid_CustomDrawCell
                AddHandler VertGrid.ValidatingEditor, AddressOf VGrid_ValidatingEditor
                AddHandler VertGrid.ShowingEditor, AddressOf VGrid_ShowingEditor
                AddHandler VertGrid.ShownEditor, AddressOf VGrid_ShownEditor
                AddHandler VertGrid.DoubleClick, AddressOf VGrid_Event_DoubleClick
                AddHandler VertGrid.CustomRecordCellEdit, AddressOf VGrid_CustomCellEditor
                AddHandler VertGrid.CustomRecordCellEditForEditing, AddressOf VGrid_CellEditorForEditing


#End Region

#Region "Spreadsheet"
            ElseIf SectionElement.Type = "Spreadsheet" Then

                ExcelModels(ModelID).ShowSpreadsheet(Me.ActiveSpreadsheet, Me)
#End Region

#Region "Label"
            ElseIf SectionElement.Type = "Label" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                LabelsCount += 1

                ReDim Preserve Labels(LabelsCount)

                Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {.Text = ActiveDataSet.DataRows(0).DataCells(0).StringValue}
                Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                'Labels(LabelsCount).ToolTipController = TTController


                TP.Controls.Add(Labels(LabelsCount))
                TP.SetCell(Labels(LabelsCount), TPRowCount, 0)
                'TP.SetColumnSpan((Labels(LabelsCount), 3)

                'TPRowCount += 1
#End Region

#Region "TextBox and Spin"


            ElseIf SectionElement.Type = "TextBox" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                Dim SCDT As New SingleCellDataTag With {
                .TargetWorksheet = Me.ActiveSpreadsheet,
                .TargetCell = ActiveDataSet.DataRows(0).DataCells(0).SourceAddress,
                .DataType = ActiveDataSet.DataColumns(0).ColumnTag.DataType,
                .IsCalculated = ActiveDataSet.DataColumns(0).ColumnTag.IsCalculated,
                .Label = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading
            }


                If ActiveDataSet.DataColumns(0).ColumnTag.MaxVal IsNot Nothing Then

                    SCDT.MaxValSet = True
                    SCDT.MaxVal = CDbl(ActiveDataSet.DataColumns(0).ColumnTag.MaxVal)

                End If

                If ActiveDataSet.DataColumns(0).ColumnTag.MinVal IsNot Nothing Then

                    SCDT.MinValSet = True
                    SCDT.MinVal = CDbl(ActiveDataSet.DataColumns(0).ColumnTag.MinVal)

                End If



                ' ReDim Preserve TextBoxes(TextEditCount)

                If SCDT.DataType = "P" And Not ActiveDataSet.RO And Not SCDT.IsCalculated Then

                    LabelsCount += 1

                    ReDim Preserve Labels(LabelsCount)

                    Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {
                    .Tag = SCDT,
                    .Text = SCDT.Label}
                    Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                    Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                    Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                    'Labels(LabelsCount).ToolTipController = TTController


                    TP.Controls.Add(Labels(LabelsCount))
                    TP.SetCell(Labels(LabelsCount), TPRowCount, 0)

                    'TPRowCount += 1
                    'TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                    'TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                    SpinEditCount += 1
                    ReDim Preserve SpinEdits(SpinEditCount)
                    SpinEdits(SpinEditCount) = New AbovoDESpinEdit With {
                                                .Name = "SpinEdit_" & SpinEditCount.ToString,
                                                        .Tag = SCDT
                                            }

                    'SpinEdTE.Properties.Increment
                    SpinEdits(SpinEditCount).Properties.Increment = 0.0025
                    SpinEdits(SpinEditCount).EnterMoveNextControl = True
                    SpinEdits(SpinEditCount).EditValue = ActiveDataSet.DataRows(0).DataCells(0).RealValue
                    SpinEdits(SpinEditCount).Properties.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                    SpinEdits(SpinEditCount).Properties.MaskSettings.Set("mask", "p5")
                    SpinEdits(SpinEditCount).Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                    SpinEdits(SpinEditCount).Properties.DisplayFormat.FormatString = "p2"


                    TP.Controls.Add(SpinEdits(SpinEditCount))
                    TP.SetCell(SpinEdits(SpinEditCount), TPRowCount, 1)

                    SpinEdits(SpinEditCount).Properties.Appearance.BackColor = Me.ActiveSpreadsheet.Cells(ActiveDataSet.DataRows(0).DataCells(0).SourceAddress).Fill.BackgroundColor

                    SpinEdits(SpinEditCount).Properties.Appearance.ForeColor = Me.ActiveSpreadsheet.Cells(ActiveDataSet.DataRows(0).DataCells(0).SourceAddress).Font.Color

                    SectionControlsCumlHeight += SpinEdits(SpinEditCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                    If ActiveDataSet.RO Or SCDT.IsCalculated = True Then

                        RefreshableControlsCount += 1

                        ReDim Preserve RefreshableControls(RefreshableControlsCount)
                        RefreshableControls(RefreshableControlsCount) = SpinEdits(SpinEditCount)

                    Else

                        AddHandler SpinEdits(SpinEditCount).Validating, AddressOf SingleCellControlValidatingEditor
                        AddHandler SpinEdits(SpinEditCount).EditValueChanged, AddressOf SingleCellDirtyMarker
                        AddHandler SpinEdits(SpinEditCount).Leave, AddressOf SingleCell_Value_Push

                    End If

                    LastBottom = SpinEdits(SpinEditCount).Bottom

                Else
                    LabelsCount += 1

                    ReDim Preserve Labels(LabelsCount)

                    Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {
                    .Tag = SCDT,
                    .Text = SCDT.Label}
                    Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                    Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                    Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                    'Labels(LabelsCount).ToolTipController = TTController


                    TP.Controls.Add(Labels(LabelsCount))
                    TP.SetCell(Labels(LabelsCount), TPRowCount, 0)


                    TextEditCount += 1
                    ReDim Preserve TextBoxes(TextEditCount)

                    TextBoxes(TextEditCount) = New AbovoDETextEdit With {
                                    .Top = LastBottom + 20,
                                    .Width = 400,
                                    .Name = "TextBox_" & TextEditCount.ToString,
                                    .ModelID = ModelID,
                                    .Tag = SCDT,
                                    .EnterMoveNextControl = True
                                }

                    TextBoxes(TextEditCount).Initialise()



                    If SCDT.TipText IsNot Nothing Then

                        'NewTE.ToolTipController = TTController
                        'Labels(LabelsCount).ToolTipController = TTController

                        Labels(LabelsCount).ToolTip = TrimSpaces(SCDT.TipText)
                        TextBoxes(TextEditCount).ToolTip = TrimSpaces(SCDT.TipText)

                    End If

                    TP.Controls.Add(TextBoxes(TextEditCount))
                    TP.SetCell(TextBoxes(TextEditCount), TPRowCount, 1)

                    SectionControlsCumlHeight += TextBoxes(TextEditCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                    If ActiveDataSet.RO Or SCDT.IsCalculated = True Then

                        RefreshableControlsCount += 1

                        ReDim Preserve RefreshableControls(RefreshableControlsCount)
                        RefreshableControls(RefreshableControlsCount) = TextBoxes(TextEditCount)

                    Else

                        AddHandler TextBoxes(TextEditCount).Validating, AddressOf SingleCellControlValidatingEditor
                        AddHandler TextBoxes(TextEditCount).EditValueChanged, AddressOf SingleCellDirtyMarker
                        AddHandler TextBoxes(TextEditCount).Leave, AddressOf SingleCell_Value_Push

                    End If

                    LastBottom = TextBoxes(TextEditCount).Bottom

                End If


#End Region

#Region "ComboBox"


            ElseIf SectionElement.Type = "ComboBox" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                Dim SCDT As New SingleCellDataTag With {
                .TargetWorksheet = Me.ActiveSpreadsheet,
                .TargetCell = ActiveDataSet.DataRows(0).DataCells(0).SourceAddress,
                .DataType = ActiveDataSet.DataColumns(0).ColumnTag.DataType,
                .IsCalculated = ActiveDataSet.DataColumns(0).ColumnTag.IsCalculated,
                .Label = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading
            }

                LabelsCount += 1
                ReDim Preserve Labels(LabelsCount)
                Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {.Text = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading}

                Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                Labels(LabelsCount).AutoSizeInLayoutControl = True
                Labels(LabelsCount).AllowHtmlString = True
                Labels(LabelsCount).Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                Labels(LabelsCount).Appearance.Options.UseTextOptions = True

                TP.Controls.Add(Labels(LabelsCount))
                TP.SetCell(Labels(LabelsCount), TPRowCount, 0)
                SectionControlsCumlHeight += Labels(LabelsCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                'TPRowCount += 1
                'TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                'TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                If Not IsNothing(ActiveDataSet.DataColumns(0).ColumnTag.RepositaryID) Then

                    'Dim ResCollection As New RepositaryItems

                    Dim EditControl As RepositaryItems.AbovoRespositaryItem
                    EditControl = RepositaryItems.GetEditor(ActiveDataSet.DataColumns(0).ColumnTag.RepositaryID, ModelID)

                    If EditControl.RepType = "CMB" Then

                        CombosCount += 1
                        ReDim Preserve Combos(CombosCount)

                        Combos(CombosCount) = EditControl.RetCombo.CreateEditor

                        Combos(CombosCount).EditValue = ActiveDataSet.DataRows(0).DataCells(0).StringValue
                        Combos(CombosCount).Properties.Appearance.Options.UseBackColor = True
                        Combos(CombosCount).Properties.Appearance.Options.UseForeColor = True

                        Combos(CombosCount).Tag = SCDT
                        Combos(CombosCount).Properties.Items.AddRange(EditControl.ListItems)
                        Combos(CombosCount).Name = "Combo_" & CombosCount.ToString
                        Combos(CombosCount).Properties.Appearance.ForeColor = Color.White
                        Combos(CombosCount).Properties.Appearance.BackColor = Me.ActiveSpreadsheet.Cells(ActiveDataSet.DataRows(0).DataCells(0).SourceAddress).Fill.BackgroundColor

                        Combos(CombosCount).EnterMoveNextControl = True

                        SectionControlsCumlHeight += Combos(CombosCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                        TP.Controls.Add(Combos(CombosCount))
                        TP.SetCell(Combos(CombosCount), TPRowCount, 1)

                        If ActiveDataSet.DataColumns(0).ColumnTag.TipText IsNot Nothing Then

                            'Combos(CombosCount).ToolTipController = TTController
                            'Labels(LabelsCount).ToolTipController = TTController

                            Labels(LabelsCount).ToolTip = TrimSpaces(SCDT.TipText)
                            Combos(CombosCount).ToolTip = TrimSpaces(SCDT.TipText)

                        End If

                        AddHandler Combos(CombosCount).Enter, AddressOf ComboOpen
                        AddHandler Combos(CombosCount).EditValueChanged, AddressOf SingleCell_Value_Push

                    End If

                End If


#End Region

#Region "DateBox"


            ElseIf SectionElement.Type = "DateBox" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                LabelsCount += 1
                ReDim Preserve Labels(LabelsCount)

                Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {.Text = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading}
                Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical

                TP.Controls.Add(Labels(LabelsCount))
                TP.SetCell(Labels(LabelsCount), TPRowCount, 0)

                'TPRowCount += 1
                'TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                'TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                DateEditCount += 1
                ReDim Preserve DateBoxes(DateEditCount)

                Dim SCDT As New SingleCellDataTag With {
                .TargetWorksheet = Me.ActiveSpreadsheet,
                .TargetCell = ActiveDataSet.DataRows(0).DataCells(0).SourceAddress,
                .DataType = ActiveDataSet.DataColumns(0).ColumnTag.DataType,
                .IsCalculated = ActiveDataSet.DataColumns(0).ColumnTag.IsCalculated,
                .Label = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading
            }

                DateBoxes(DateEditCount) = New AbovoDEDateEdit With {
                    .EnterMoveNextControl = True,
                    .TargetCell = SCDT.TargetCell,
                    .ModelID = ModelID,
                    .TargetWorksheet = Me.ActiveSpreadsheet,
                    .Tag = SCDT
                }

                DateBoxes(DateEditCount).Initialise()

                'DateBoxes(DateEditCount).DateTime = DateTime.FromOADate(ActiveDataSet.DataRows(0).DataCells(0).IntValue)

                DateBoxes(DateEditCount).BackColor = AbovoBlue
                DateBoxes(DateEditCount).ForeColor = Color.White

                If SCDT.TipText IsNot Nothing Then

                    'DateBoxes(DateEditCount).ToolTipController = TTController
                    'Labels(LabelsCount).ToolTipController = TTController

                    Labels(LabelsCount).ToolTip = TrimSpaces(SCDT.TipText)
                    DateBoxes(DateEditCount).ToolTip = TrimSpaces(SCDT.TipText)

                End If

                If SCDT.DataType = "DM" Then

                    DateBoxes(DateEditCount).Properties.DisplayFormat.FormatString = "dd-MMM"

                End If

                AddHandler DateBoxes(DateEditCount).EditValueChanged, AddressOf SingleCell_Value_Push

                SectionControlsCumlHeight += DateBoxes(DateEditCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                TP.Controls.Add(DateBoxes(DateEditCount))
                TP.SetCell(DateBoxes(DateEditCount), TPRowCount, 1)
#End Region

#Region "Spacer"

            ElseIf SectionElement.Type = "Spacer" Then

                'LabelsCount += 1
                'ReDim Preserve Labels(LabelsCount)
                'Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl
                'Labels(LabelsCount).Visible = False
                'Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                'TP.Controls.Add(Labels(LabelsCount))
                'TP.SetCell(Labels(LabelsCount), TPRowCount, 0)
                'SectionControlsCumlHeight += Labels(LabelsCount).Height + (2 * DefaultTablePanelPadding)
#End Region

#Region "Link"

            ElseIf SectionElement.Type = "Link" Then

                For Each LControl In SectionElement.LinkControls

                    If LControl.LinkType = "GOTO" Then

                        Dim ButtText As String = LControl.LinkText

                        ButtText = ButtText.Replace("vblf", "<br>")
                        Dim NewButton As New DevExpress.XtraEditors.SimpleButton

                        NewButton.AllowHtmlDraw = DefaultBoolean.True
                        NewButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                        NewButton.Text = ButtText
                        NewButton.ToolTip = LControl.LinkTip

                        LControl.LinkReturnID = GSID

                        NewButton.Tag = LControl

                        NewButton.Padding = New System.Windows.Forms.Padding(10)
                        Dim bestSize As Size = NewButton.CalcBestSize()
                        NewButton.Width = bestSize.Width
                        NewButton.Height = bestSize.Height

                        AddHandler NewButton.Click, AddressOf Link_ButtonClick
                        TP.Controls.Add(NewButton)
                        TP.SetCell(NewButton, TPRowCount, 0)
                        SectionControlsCumlHeight += NewButton.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                        TPRowCount += 1
                        TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                        TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                    End If

                Next



#End Region

#Region "TextArea"

            ElseIf SectionElement.Type = "TextArea" Then

                Dim NewRTF As New DevExpress.XtraEditors.LabelControl

                NewRTF.AllowHtmlString = True
                NewRTF.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                NewRTF.Appearance.Options.UseTextOptions = True
                NewRTF.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Vertical
                'NewRTF.HyperlinkClick += LabelControl1_HyperlinkClick

                NewRTF.Text = SectionElement.Description
                TP.Controls.Add(NewRTF)
                TP.SetCell(NewRTF, TPRowCount, 0)
                SectionControlsCumlHeight += NewRTF.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

#End Region

#Region "Browser"

            ElseIf SectionElement.Type = "Browser" Then

                Dim NewWB As New System.Windows.Forms.WebBrowser

                NewWB.DocumentText = SectionElement.Description

                TP.Controls.Add(NewWB)
                TP.SetCell(NewWB, TPRowCount, 0)
                TP.SetColumnSpan(NewWB, 3)
                SectionControlsCumlHeight += NewWB.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

#End Region

#Region "CompundMappedTable"

            ElseIf SectionElement.Type = "CompoundMappedTable" Then


                Dim DockPosition As DockStyle = DockStyle.Right
                Dim x As Integer = 0
                Dim RowCount As Integer = 0
                Dim ColCount As Integer = 0
                Dim LastRowIndex As Integer = 0
                Dim ConPadding As Padding
                ConPadding.All = 10
                ConPadding.Top = 5
                ConPadding.Bottom = 5
                TP.Rows.Clear()
                TP.Columns.Clear()

                TP.Width = Me.Width

                Dim Newpadding As New Padding()
                Newpadding.Right += 20
                TP.Padding = DefaultTablePanelPadding

                Dim CMTab As CompoundMappedTable = SectionElement.CompoundTable

                'Define the CMT starting position and size

                Dim StartIndexCol As Integer = 0
                Dim StartIndexRow As Integer = 0
                Dim TotalColumnCount As Integer = 0

                Dim TestRange As DevExpress.Spreadsheet.CellRange = Nothing

                If Microsoft.VisualBasic.Left(CMTab.StartingCol, 4) = "Col-" Then

                    StartIndexCol = CInt(Mid(CMTab.StartingCol, 5))

                Else

                    Dim StartRnge As DevExpress.Spreadsheet.CellRange = Me.ActiveWorkbook.Range(Mid(CMTab.StartingCol, 5))
                    StartIndexCol = StartRnge.LeftColumnIndex

                End If

                If Microsoft.VisualBasic.Left(CMTab.StartingRow, 4) = "Row-" Then

                    StartIndexRow = CInt(Mid(CMTab.StartingRow, 5))

                Else

                    TestRange = Me.ActiveWorkbook.Range(Mid(CMTab.StartingRow, 4))
                    StartIndexRow = TestRange.TopRowIndex

                End If

                For Each CMTColDef As CompoundMappedTableColumnDefinition In CMTab.CMTColDefs

                    If Microsoft.VisualBasic.Left(CMTColDef.ColSetType, 6) = "Fixed-" Then

                        CMTColDef.ColCount = CInt(Mid(CMTColDef.ColSetType, 7))
                        TotalColumnCount += CInt(Mid(CMTColDef.ColSetType, 7))

                    Else

                        TestRange = Me.ActiveWorkbook.Range(Mid(CMTColDef.RepeatsBy, 6))
                        CMTColDef.ColCount = TestRange.ColumnCount
                        CMTColDef.StartColIndex = TestRange.LeftColumnIndex

                        For x = LastRowIndex To (LastRowIndex + CMTColDef.ColCount - 1)

                            TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.AutoSize, 50))

                            TP.Columns(x).Tag = New MappedTableColumnTag With {
                            .ColDef = CMTColDef.ColSetType,
                            .ColSetID = CMTColDef.ColSetID,
                            .Index = x
                        }

                        Next

                        TotalColumnCount += CMTColDef.ColCount

                    End If

                    If Not IsNothing(CMTColDef.RepeatsBy) Then CMTColDef.ElementRepeats = True



                Next

                Dim RowRepeats As Integer = 0
                Dim ElementRepeats As Integer = 0
                Dim OverallRowIndex As Integer = 0

                For Each MTRow As MappedTableRow In CMTab.MappedTableRows

                    RowRepeats = 0

                    If MTRow.RepeatsBy IsNot Nothing Then

                        If Microsoft.VisualBasic.Left(MTRow.RepeatsBy, 4) = "Row-" Then

                            RowRepeats = CInt(Mid(MTRow.RepeatsBy, 5))

                        Else

                            TestRange = Me.ActiveWorkbook.Range(Mid(MTRow.RepeatsBy, 5))
                            RowRepeats = CInt(MTRow.RepeatsBy)

                        End If

                    End If

                    Dim CurrentRowIndex As Integer = 0
                    Dim CurrColDef As Integer = 0

                    For x = 0 To RowRepeats

                        If CMTab.CMTColDefs(CurrColDef).ElementRepeats = True Then



                        End If

                        CurrColDef += 1

                        MTRow.RowIndex = OverallRowIndex

                        OverallRowIndex += 1

                    Next


                Next


#End Region

#Region "MappedTable"

            ElseIf SectionElement.Type = "MappedTable" Then

                Dim MTab As MappedTable = SectionElement.MappedTableSection
                Dim DockPosition As DockStyle = DockStyle.Right
                Dim x As Integer = 0
                Dim RowCount As Integer = CInt(MTab.NumRows)
                Dim ColCount As Integer = CInt(MTab.NumCols)
                Dim ConPadding = DefaultTablePanelPadding
                ConPadding.All = 10
                ConPadding.Top = 5
                ConPadding.Bottom = 5
                TP.Rows.Clear()
                TP.Columns.Clear()

                TP.Width = Me.Width

                'Dim Newpadding = DefaultTablePanelPadding
                'Newpadding.Right += 20
                TP.Padding = DefaultTablePanelPadding

                Dim MaxHeight As Integer = 0

                'Dim lastItem As LayoutControlItem = Nothing
                Dim DoneFirst As Boolean = False
                Dim CurrentCol As Integer = 0
                Dim CurrentRow As Integer = -1

                For x = 1 To ColCount + 1

                    TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.AutoSize, TP.Width / CInt(ColCount)))
                    TP.Columns(CurrentCol).Tag = x.ToString

                    CurrentCol += 1

                Next



                For Each MappedTableRow In MTab.MappedTableRows

                    TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                    CurrentRow += 1
                    CurrentCol = 0

                    Dim LastRight As Integer = 0
                    Dim ColSpan As Integer = 1
                    Dim RowSpan As Integer = 1

                    Dim CurrControl As Control = Nothing

                    For Each MTE As MappedTableElement In MappedTableRow.MappedTableElements

                        ColSpan = 1
                        RowSpan = 1

                        DockPosition = DockStyle.Right

                        If MTE.Type = "Skip" Then

                            GoTo NextCell

                        ElseIf MTE.Type = "Link" Then

                            ADEHyperlinksCount += 1
                            ReDim Preserve ADEHyperlinks(ADEHyperlinksCount)

                            Dim TargetInterface As String = Me.ActiveSpreadsheet.Cells(MTE.TELinks(0).CSSource).DisplayText
                            Dim LinkTag As New ElementInterfaceLinkTag With {
                            .LinkReturnID = CSID,
                            .LinkReturnName = DataPres.PresName,
                            .LinkData = TargetInterface,
                            .LinkReturnGroup = GSID,
                            .LinkTargetSection = MTE.TELinks(0).LinkTargetSection
                        }

                            ADEHyperlinks(ADEHyperlinksCount) = New AbovoDEHyperlinkLabel With {
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .AutoSizeHorizontal = True,
                            .TargetCell = MTE.TELinks(0).CSSource,
                            .Dock = DockStyle.Fill,
                            .Tag = LinkTag
                        }

                            DockPosition = DockStyle.Left

                            ADEHyperlinks(ADEHyperlinksCount).Initialise()
                            CurrControl = ADEHyperlinks(ADEHyperlinksCount)
                            AddHandler ADEHyperlinks(ADEHyperlinksCount).HyperlinkClick, AddressOf DEHyperLinkLabel_linkClick

                        ElseIf MTE.Type = "ComboBox" Then


                            Dim TECB As TEComboBox = MTE.TEComboBoxes(0)
                            Dim SCDT As New SingleCellDataTag With {
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .DataType = TECB.DataType,
                            .TargetCell = TECB.CSSource
                        }

                            ADEComboBoxesCount += 1
                            ReDim Preserve ADEComboBoxes(ADEComboBoxesCount)
                            ADEComboBoxes(ADEComboBoxesCount) = New AbovoDEComboBox With {
                            .Tag = SCDT,
                            .RepID = TECB.RepID,
                            .Dock = DockStyle.Fill,
                            .AutoSizeHorizontal = True
                        }

                            RefreshableControlsCount += 1
                            ReDim Preserve RefreshableControls(RefreshableControlsCount)
                            RefreshableControls(RefreshableControlsCount) = ADEComboBoxes(ADEComboBoxesCount)

                            ADEComboBoxes(ADEComboBoxesCount).InitialiseStandard(TECB.RepID)
                            AddHandler ADEComboBoxes(ADEComboBoxesCount).EditValueChanged, AddressOf SingleCell_Value_Push
                            CurrControl = ADEComboBoxes(ADEComboBoxesCount)

                        ElseIf MTE.Type = "Label" Then

                            Dim TEL As TELabel = MTE.TELabels(0)

                            Dim SCDT As New SingleCellDataTag With {
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .TargetCell = TEL.CSSource
                        }

                            Dim DoTitle As Integer = 0

                            If TEL.IsTitle Is Nothing OrElse TEL.IsTitle = "False" Then

                                DoTitle = 0

                            ElseIf TEL.IsTitle = "True" Then

                                DoTitle = 1

                            Else

                                DoTitle = CInt(TEL.IsTitle)

                            End If

                            ADELabelsCount += 1
                            ReDim Preserve ADELabels(ADELabelsCount)

                            If Not IsNothing(MTE.ColSpan) Then

                                ColSpan = CInt(MTE.ColSpan)

                            End If

                            If Not IsNothing(MTE.RowSpan) Then

                                RowSpan = CInt(MTE.RowSpan)

                            End If

                            ADELabels(ADELabelsCount) = New AbovoDELabel With {
                            .Tag = SCDT,
                            .AutoSizeHorizontal = True,
                            .Dock = DockStyle.Fill,
                            .IsBold = IIf(TEL.IsBold = "True", True, False),
                            .IsUnderline = IIf(TEL.IsUnderline = "True", True, False),
                            .IsTitle = DoTitle
                        }
                            ADELabels(ADELabelsCount).AutoSizeInLayoutControl = True
                            'ADELabels(ADELabelsCount).RealAutoSizeMode = AutoSizeMode.Vertical
                            If TEL.Alignment IsNot Nothing Then

                                Select Case TEL.Alignment.ToUpper
                                    Case "NEAR", "LEFT"
                                        ADELabels(ADELabelsCount).TxtHlignment = DevExpress.Utils.HorzAlignment.Near
                                    Case "FAR", "RIGHT"
                                        ADELabels(ADELabelsCount).TxtHlignment = DevExpress.Utils.HorzAlignment.Far
                                    Case "CENTER", "CENTRE", "MIDDLE"
                                        ADELabels(ADELabelsCount).TxtHlignment = DevExpress.Utils.HorzAlignment.Center
                                End Select

                            Else

                                ADELabels(ADELabelsCount).TxtHlignment = DevExpress.Utils.HorzAlignment.Far

                            End If

                            If TEL.IsBordered = "True" Then

                                ADELabels(ADELabelsCount).DoBorder = True

                            End If

                            If TEL.IsStatic IsNot "True" Then

                                RefreshableControlsCount += 1
                                ReDim Preserve RefreshableControls(RefreshableControlsCount)
                                RefreshableControls(RefreshableControlsCount) = ADELabels(ADELabelsCount)

                            Else

                                ADELabels(ADELabelsCount).IsStatic = True

                            End If

                            ADELabels(ADELabelsCount).Initialise()

                            If ADELabels(ADELabelsCount).TxtHlignment = HorzAlignment.Far Then
                                DockPosition = DockStyle.Right
                            Else
                                DockPosition = DockStyle.Left
                            End If

                            CurrControl = ADELabels(ADELabelsCount)

                        ElseIf MTE.Type = "DateEdit" Then

                            ADEDateEditsCount += 1
                            ReDim Preserve ADEDateEdits(ADEDateEditsCount)

                            If Not IsNothing(MTE.ColSpan) Then

                                ColSpan = CInt(MTE.ColSpan)

                            End If

                            If Not IsNothing(MTE.RowSpan) Then

                                RowSpan = CInt(MTE.RowSpan)

                            End If

                            ADEDateEdits(ADEDateEditsCount) = New AbovoDEDateEdit With {
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .Dock = DockStyle.Fill,
                            .TargetCell = MTE.TELabels(0).CSSource
                        }

                            If MTE.TEDateEdits(0).IsStatic = "True" Then

                                ADEDateEdits(ADEDateEditsCount).IsStatic = True

                            Else

                                ADEDateEdits(ADEDateEditsCount).IsStatic = False
                                RefreshableControlsCount += 1
                                ReDim Preserve RefreshableControls(RefreshableControlsCount)
                                RefreshableControls(RefreshableControlsCount) = ADEDateEdits(ADEDateEditsCount)

                            End If

                        ElseIf MTE.Type = "TextEdit" Then

                            ADETextEditsCount += 1
                            ReDim Preserve ADETextEdits(ADETextEditsCount)

                            If Not IsNothing(MTE.ColSpan) Then

                                ColSpan = CInt(MTE.ColSpan)

                            End If

                            If Not IsNothing(MTE.RowSpan) Then

                                RowSpan = CInt(MTE.RowSpan)

                            End If

                            Dim TETE As TETextBox = MTE.TETextBoxes(0)

                            Dim SCDataDag As New SingleCellDataTag With {
                            .DataType = TETE.DataType,
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .TargetCell = TETE.CSSource
                        }

                            If TETE.MaxVal IsNot Nothing Then

                                SCDataDag.MaxValSet = True
                                SCDataDag.MaxVal = CDbl(TETE.MaxVal)

                            End If

                            If TETE.MinVal IsNot Nothing Then

                                SCDataDag.MinValSet = True
                                SCDataDag.MinVal = CDbl(TETE.MinVal)

                            End If

                            If MTE.ToolTip IsNot Nothing Then

                                ADETextEdits(ADETextEditsCount).ToolTip = TrimSpaces(MTE.ToolTip)

                            End If

                            ADETextEdits(ADETextEditsCount) = New AbovoDETextEdit With {
                                .ModelID = ModelID,
                                .EnterMoveNextControl = True,
                                .Dock = DockStyle.Fill,
                            .Tag = SCDataDag
                        }

                            If MTE.TETextBoxes(0).IsReadOnly = "True" Then

                                ADETextEdits(ADETextEditsCount).IsReadOnly = True

                            Else

                                ADETextEdits(ADETextEditsCount).IsReadOnly = False
                                RefreshableControlsCount += 1
                                ReDim Preserve RefreshableControls(RefreshableControlsCount)
                                RefreshableControls(RefreshableControlsCount) = ADETextEdits(ADETextEditsCount)

                            End If

                            ADETextEdits(ADETextEditsCount).Initialise()

                            If ADETextEdits(ADETextEditsCount).Properties.Appearance.TextOptions.HAlignment = HorzAlignment.Far Then
                                DockPosition = DockStyle.Right
                            Else
                                DockPosition = DockStyle.Left
                            End If


                            CurrControl = ADETextEdits(ADETextEditsCount)
                            ADETextEdits(ADETextEditsCount).Properties.EditValueChangedFiringMode = EditValueChangedFiringMode.Buffered
                            ADETextEdits(ADETextEditsCount).Properties.EditValueChangedDelay = 400

                            AddHandler ADETextEdits(ADETextEditsCount).Validating, AddressOf SingleCellControlValidatingEditor
                            AddHandler ADETextEdits(ADETextEditsCount).EditValueChanged, AddressOf SingleCellDirtyMarker
                            AddHandler ADETextEdits(ADETextEditsCount).Leave, AddressOf SingleCell_Value_Push

                        ElseIf MTE.Type = "Hyperlink" Then

                            ADETextEditsCount += 1
                            ReDim Preserve ADETextEdits(ADETextEditsCount)

                            If Not IsNothing(MTE.ColSpan) Then

                                ColSpan = CInt(MTE.ColSpan)

                            End If

                            If Not IsNothing(MTE.RowSpan) Then

                                RowSpan = CInt(MTE.RowSpan)

                            End If

                            Dim TETE As TETextBox = MTE.TETextBoxes(0)

                            Dim SCDataDag As New SingleCellDataTag With {
                            .DataType = TETE.DataType,
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .TargetCell = TETE.CSSource
                        }

                            If TETE.MaxVal IsNot Nothing Then

                                SCDataDag.MaxValSet = True
                                SCDataDag.MaxVal = CDbl(TETE.MaxVal)

                            End If

                            If TETE.MinVal IsNot Nothing Then

                                SCDataDag.MinValSet = True
                                SCDataDag.MinVal = CDbl(TETE.MinVal)

                            End If

                            If MTE.ToolTip IsNot Nothing Then

                                ADETextEdits(ADETextEditsCount).ToolTip = TrimSpaces(MTE.ToolTip)

                            End If

                            ADETextEdits(ADETextEditsCount) = New AbovoDETextEdit With {
                                .ModelID = ModelID,
                                .EnterMoveNextControl = True,
                                .Dock = DockStyle.Fill,
                            .Tag = SCDataDag
                        }

                            If MTE.TETextBoxes(0).IsReadOnly = "True" Then

                                ADETextEdits(ADETextEditsCount).IsReadOnly = True

                            Else

                                ADETextEdits(ADETextEditsCount).IsReadOnly = False
                                RefreshableControlsCount += 1
                                ReDim Preserve RefreshableControls(RefreshableControlsCount)
                                RefreshableControls(RefreshableControlsCount) = ADETextEdits(ADETextEditsCount)

                            End If

                            ADETextEdits(ADETextEditsCount).Initialise()

                            If ADETextEdits(ADETextEditsCount).Properties.Appearance.TextOptions.HAlignment = HorzAlignment.Far Then
                                DockPosition = DockStyle.Right
                            Else
                                DockPosition = DockStyle.Left
                            End If


                            CurrControl = ADETextEdits(ADETextEditsCount)
                            ADETextEdits(ADETextEditsCount).Properties.EditValueChangedFiringMode = EditValueChangedFiringMode.Buffered
                            ADETextEdits(ADETextEditsCount).Properties.EditValueChangedDelay = 400

                            AddHandler ADETextEdits(ADETextEditsCount).Validating, AddressOf SingleCellControlValidatingEditor
                            AddHandler ADETextEdits(ADETextEditsCount).EditValueChanged, AddressOf SingleCellDirtyMarker
                            AddHandler ADETextEdits(ADETextEditsCount).Leave, AddressOf SingleCell_Value_Push

                        End If


                        If ColSpan > 1 Then

                            TP.SetColumnSpan(CurrControl, ColSpan)

                        End If

                        If RowSpan = 3 Then

                            TP.SetRowSpan(CurrControl, RowSpan)

                        End If

                        CurrControl.Margin = ConPadding

                        If MaxHeight < CurrControl.Height Then MaxHeight = CurrControl.Height

                        CurrControl.Dock = DockPosition

                        TP.Controls.Add(CurrControl)
                        TP.SetCell(CurrControl, CurrentRow, CurrentCol)
                        'TP.SetColumnSpan(CurrControlGridControls(GridCount), 3)
NextCell:


                        CurrentCol += 1
                        CurrentCol += ColSpan - 1

                        CurrControl = Nothing

                    Next

                    SectionControlsCumlHeight += MaxHeight + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                    MaxHeight = 0

                Next

#End Region

#Region "Control Group"

            ElseIf SectionElement.Type = "ControlGroup" Then

                Dim LastRight As Integer = 0

                'For Each GridButtControl In SectionElement.GridControls

                '    ButtText = GridButtControl.CommandText.Replace("vblf", "<br>")

                '    NewButton.AllowHtmlDraw = DefaultBoolean.True
                '    NewButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                '    NewButton.Text = ButtText
                '    NewButton.ToolTip = GridButtControl.CommandTip

                '    GridButtControl.AttachedGrid = GridControls(GridCount)

                '    NewButton.Tag = GridButtControl



                '    'NewButton.Padding = New Padding(10)
                '    Dim bestSize As Size = NewButton.CalcBestSize()
                '    NewButton.Width = bestSize.Width
                '    NewButton.Height = bestSize.Height

                '    AddHandler NewButton.Click, AddressOf Grid_ButtonClick
                '    'TP.Controls.Add(NewButton)
                '    'TP.SetCell(NewButton, TPRowCount, 0)
                '    'NewButton.Left = LastRight + 10
                '    'LastRight = NewButton.Right
                '    itemButton.Control = NewButton


                'Next

                Dim NewButton As DevExpress.XtraEditors.SimpleButton = Nothing
                Dim ButtText As String = ""

                Dim Layout As LayoutControl = New LayoutControl()
                Layout.BackColor = Color.White

                Layout.OptionsView.EnableTransparentBackColor = False
                Layout.Padding = DefaultTablePanelPadding

                Layout.Dock = DockStyle.Fill

                Layout.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
                Layout.LookAndFeel.UseDefaultLookAndFeel = False
                Layout.OptionsView.EnableTransparentBackColor = True
                Layout.Root.AppearanceGroup.BackColor = System.Drawing.Color.Transparent
                Layout.Root.AppearanceGroup.Options.UseBackColor = True

                TP.Controls.Add(Layout)
                TP.SetCell(Layout, TPRowCount, 0)

                Dim groupButtons As LayoutControlGroup = Layout.Root.AddGroup()
                groupButtons.Name = "GroupButtons"
                groupButtons.GroupBordersVisible = False

                Dim lastItem As LayoutControlItem = Nothing
                Dim DoneFirst As Boolean = False

                For Each IControl In SectionElement.InterfaceControls

                    NewButton = New DevExpress.XtraEditors.SimpleButton()


                    Dim itemButton As LayoutControlItem = Nothing

                    If Not DoneFirst Then

                        itemButton = Layout.AddItem(groupButtons, InsertType.Left)

                        DoneFirst = True
                        lastItem = itemButton

                    Else

                        itemButton = Layout.AddItem(lastItem, InsertType.Right)
                        DoneFirst = True
                        lastItem = itemButton

                    End If

                    ButtText = IControl.ItemTxt

                    ButtText = ButtText.Replace("vblf", "<br>")

                    NewButton.AllowHtmlDraw = DefaultBoolean.True
                    NewButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                    NewButton.Text = ButtText
                    NewButton.ToolTip = IControl.ItemTip
                    NewButton.Tag = IControl.ItemData

                    NewButton.Left = LastRight + 10

                    NewButton.Padding = New System.Windows.Forms.Padding(10)
                    Dim bestSize As Size = NewButton.CalcBestSize()
                    NewButton.Width = bestSize.Width
                    NewButton.Height = bestSize.Height
                    itemButton.Control = NewButton
                    AddHandler NewButton.Click, AddressOf Interface_ButtonClick

                Next

                Dim emptySpace As New EmptySpaceItem

                emptySpace.Parent = groupButtons


                SectionControlsCumlHeight += NewButton.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                TPRowCount += 1

                TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))

                TP.Rows(TPRowCount).Tag = TPRowCount.ToString

#End Region

            End If

            'AcContainers(AcContainersCount).Height = SectionControlsCumlHeight
            'AcContainers(AcContainersCount).Width = Me.Width
            'AcContainers(AcContainersCount).Appearance.BackColor = AbovoBlue

            'Add a small spacer
            TPRowCount += 1
            TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.Absolute, 50))
            TP.Rows(TPRowCount).Tag = TPRowCount.ToString
            'LabelsCount += 1
            'ReDim Preserve Labels(LabelsCount)
            'Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl
            'Labels(LabelsCount).Visible = False
            'Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
            'TP.Controls.Add(Labels(LabelsCount))
            'TP.SetCell(Labels(LabelsCount), TPRowCount, 0)
            'SectionControlsCumlHeight += Labels(LabelsCount).Height + (2 * DefaultTablePanelPadding)

        Next

        'Do not assign TP.Height from SectionControlsCumlHeight here.
        '
        'Several controls (especially VGrid) are deliberately resized again after
        'the real tab-page viewport is known. A fixed height captured during the
        'build would leave the XtraTabPage AutoScroll extent stale and can make
        'the bottom of a tall VGrid unreachable.
        '
        'With AutoSize + DockStyle.Top, TablePanel derives its height from the
        'actual final row heights and the XtraTabPage can scroll to the true end.
        TP.EndInit()
        TP.PerformLayout()

        Formatter.FormatTablePanel(TP)

#If DEBUG Then
#End If

    End Sub

#Region "Public Methods"

    Private Sub RebuildSelectedSectionAtomically(ByVal SectionIndex As Integer)

        If SectionIndex < 0 OrElse SectionIndex >= DataPres.Sections.Length Then Return
        If XtraTabControlNewGIT.SelectedTabPageIndex <> SectionIndex Then
            RebuildSection(SectionIndex)
            Return
        End If

        Dim LocalUpdateHeld As Boolean = False

        'Use the same native redraw suppression that removed the first-time lazy
        'tab jump. While the selected tab is disposed/recreated, Windows keeps the
        'last complete frame on screen. Painting is re-enabled only after the new
        'controls have been built and given their final viewport layout.
        SuspendLazyTabTransitionRedraw()

        Try

            XtraTabControlNewGIT.BeginUpdate()
            LocalUpdateHeld = True

            RebuildSection(SectionIndex)

            'Unlike a first-time lazy build, this page is already selected, so its
            'real viewport is immediately available. Give the replacement controls
            'one authoritative final layout pass while native painting is still off.
            XtraTabControlNewGIT.PerformLayout()

            Dim SelectedPage As XtraTabPage =
                XtraTabControlNewGIT.TabPages(SectionIndex)

            If SelectedPage IsNot Nothing AndAlso
               Not SelectedPage.IsDisposed Then

                SelectedPage.PerformLayout()

            End If

            Dim SelectedTP As TablePanel = Nothing

            If TPs IsNot Nothing AndAlso
               SectionIndex < TPs.Length Then

                SelectedTP = TPs(SectionIndex)

            End If

            If SelectedTP IsNot Nothing AndAlso
               Not SelectedTP.IsDisposed Then

                SelectedTP.PerformLayout()
                ApplySectionFontAndGridLayout(SelectedTP)
                SelectedTP.PerformLayout()

            End If

        Finally

            If LocalUpdateHeld Then

                Try
                    XtraTabControlNewGIT.EndUpdate()
                Catch
                End Try

            End If

            ResumeLazyTabTransitionRedraw()

        End Try

    End Sub

    Sub RebuildSection(SectionIndex As Integer)

        If SectionIndex < 0 OrElse SectionIndex >= DataPres.Sections.Length Then Return

        Dim State As InterfaceSectionRuntimeState = Nothing

        If SectionRuntimeStates.TryGetValue(SectionIndex, State) Then
            State.IsDirty = True
            State.NeedsPresentationRedefinition = True
        End If

        BuildSection(SectionIndex, True)

    End Sub

    Sub RebuildAllSections()

        'Retained as a safe fallback/debug path.  Normal structural changes now use
        'InterfaceDependencyManager and invalidate only worksheet-dependent sections.
        For SectionIndex As Integer = 0 To DataPres.Sections.Length - 1

            Dim State As InterfaceSectionRuntimeState = Nothing

            If SectionRuntimeStates.TryGetValue(SectionIndex, State) Then
                State.IsDirty = True
                State.NeedsPresentationRedefinition = True
            End If

            If XtraTabControlNewGIT.SelectedTabPageIndex = SectionIndex AndAlso AmActivated AndAlso Me.Visible Then
                RebuildSelectedSectionAtomically(SectionIndex)
            Else
                UnloadSection(SectionIndex)
            End If

        Next

    End Sub

    Public Sub ClearAllTPs()
        On Error Resume Next

        If TPs Is Nothing Then Return

        For Each TP In TPs

            If TP Is Nothing OrElse TP.IsDisposed Then Continue For

            For Each C As Control In TP.Controls
                If C IsNot Nothing AndAlso Not C.IsDisposed Then C.Dispose()
            Next

            TP.Controls.Clear()
            TP.Dispose()

        Next

    End Sub

    Public Sub ClearAllGrids()

        On Error Resume Next

        '------------------------------------------------------------------
        'Detach VGrid category/footer extenders FIRST.
        '
        'These objects subscribe directly to VGrid mouse/custom-draw events
        'and also own the synthetic footer rows. They must be detached while
        'the VGrid and its row hierarchy are still alive.
        '------------------------------------------------------------------
        If VGridCategoryExtenders IsNot Nothing Then

            For Each Extender As VGridCategoryButtonExtender In VGridCategoryExtenders

                If Extender IsNot Nothing Then
                    Extender.DetachForDisposal()
                End If

            Next

            VGridCategoryExtenders.Clear()

        End If


        If GridControls IsNot Nothing Then

            For Each gc In GridControls

                If gc Is Nothing Then Continue For

                RemoveHandler gc.ProcessGridKey, AddressOf GridControl_ProcessGridKey

                Dim gcSource As AbovoUnboundSource =
                    TryCast(gc.DataSource, AbovoUnboundSource)

                If gcSource IsNot Nothing Then
                    RemoveHandler gcSource.ValueNeeded, AddressOf UnboundDS_ValueNeeded
                    RemoveHandler gcSource.ValuePushed, AddressOf UnboundDS_ValuePushed
                End If

                gc.DataSource = Nothing

                Dim gcView As GridView = TryCast(gc.FocusedView, GridView)
                If gcView IsNot Nothing Then
                    gcView.Columns.Clear()
                Else
                    Dim bgvView As BandedGridView = TryCast(gc.FocusedView, BandedGridView)
                    If bgvView IsNot Nothing Then
                        bgvView.Columns.Clear()
                        bgvView.Bands.Clear()
                    End If
                End If

                gc.Dispose()
                gc = Nothing

            Next

            GridCount = -1
            GridControls = Nothing
            ReDim GridControls(-1)

        End If


        If VertGridControls IsNot Nothing Then

            For Each vg In VertGridControls

                If vg Is Nothing Then Continue For

                'Remove all handlers owned by DataInterfaceTemplate before
                'breaking the VGrid row/data-source structure.
                RemoveHandler vg.KeyDown, AddressOf VGrid_KeyDown
                RemoveHandler vg.EditorKeyDown, AddressOf VGrid_KeyDown
                RemoveHandler vg.CustomDrawRowValueCell, AddressOf VGrid_CustomDrawCell
                RemoveHandler vg.CustomDrawRowValueCell, AddressOf LiveVGrid_CustomDrawCell
                RemoveHandler vg.ValidatingEditor, AddressOf VGrid_ValidatingEditor
                RemoveHandler vg.ShowingEditor, AddressOf VGrid_ShowingEditor
                RemoveHandler vg.ShownEditor, AddressOf VGrid_ShownEditor
                RemoveHandler vg.DoubleClick, AddressOf VGrid_Event_DoubleClick
                RemoveHandler vg.CustomRecordCellEdit, AddressOf VGrid_CustomCellEditor
                RemoveHandler vg.CustomRecordCellEditForEditing, AddressOf VGrid_CellEditorForEditing

                Dim vgSource As AbovoUnboundSource =
                    TryCast(vg.DataSource, AbovoUnboundSource)

                If vgSource IsNot Nothing Then
                    RemoveHandler vgSource.ValueNeeded, AddressOf UnboundDS_ValueNeeded
                    RemoveHandler vgSource.ValuePushed, AddressOf UnboundDS_ValuePushed
                End If

                'Disconnect data first, then clear the generated/category rows.
                'At this point the category extender has already removed any
                'synthetic footer row and unsubscribed its own handlers.
                vg.DataSource = Nothing
                vg.Rows.Clear()
                vg.Dispose()

            Next

        End If

        VertGridCount = -1
        VertGridControls = Nothing
        ReDim VertGridControls(-1)


        If UnboundDataSources IsNot Nothing Then

            For Each UBS In UnboundDataSources

                If UBS Is Nothing Then Continue For

                RemoveHandler UBS.ValueNeeded, AddressOf UnboundDS_ValueNeeded
                RemoveHandler UBS.ValuePushed, AddressOf UnboundDS_ValuePushed

                UBS.Dispose()
                UBS = Nothing

            Next

            UnboundDataSources = Nothing
            UBSDataSourceCount = -1
            ReDim UnboundDataSources(-1)

        End If


        If UsedGridVIEWS IsNot Nothing Then

            For Each GV In UsedGridVIEWS

                If GV Is Nothing Then Continue For

                GV.Dispose()
                GV = Nothing

            Next

            UsedGridVIEWS = Nothing
            ReDim UsedGridVIEWS(-1)
            GridViewCount = -1

        End If


        If UsedBANDedGridVIEWS IsNot Nothing Then

            For Each BGV In UsedBANDedGridVIEWS

                If BGV Is Nothing Then Continue For

                BGV.Dispose()
                BGV = Nothing

            Next

            UsedBANDedGridVIEWS = Nothing
            ReDim UsedBANDedGridVIEWS(-1)
            BandGridViewsCount = -1

        End If

        'The band/column arrays also contain objects from the disposed views.
        'Reset them so the next section build starts with clean indexes.
        UsedBANDedGridViewBANDS = Nothing
        ReDim UsedBANDedGridViewBANDS(-1)
        BandGridViewBandsCount = -1

        UsedBANDedGridViewCOLS = Nothing
        ReDim UsedBANDedGridViewCOLS(-1)
        BandGridViewColsCount = -1

    End Sub

    Public Sub RefreshROControls()

        For Each RefreshableObject In RefreshableControls

            Try

                RefreshableObject.RefreshData()

            Catch ex As Exception


            End Try

        Next


    End Sub
    Sub SetFooterOn(SetGridView As GridView, Optional ByVal ShowTotalsWord As Boolean = True)

        FooterOn = True
        UsedGridVIEWS(GridViewCount).OptionsView.ShowFooter = True


        'If ShowTotalsWord Then

        '    Dim itemCust As New GridGroupSummaryItem
        '    itemCust.FieldName = UsedGridVIEWS(GridViewCount).Columns(0).FieldName
        '    itemCust.SummaryType = DevExpress.Data.SummaryItemType.Custom
        '    itemCust.DisplayFormat = "Totals"
        '    itemCust.ShowInGroupColumnFooter = UsedGridVIEWS(GridViewCount).Columns(0)
        '    UsedGridVIEWS(GridViewCount).GroupSummary.Add(itemCust)

        'End If

        FooterDone = True
        SetGridView.FooterPanelHeight += (2 * DefaultGridCellPadding)

    End Sub

    Sub SetBandedFooterOn(SetGridView As GridView)


    End Sub


#End Region

#Region "Menu Button Actions"
    Sub AddMenuInterface(InterfaceMode As String)

        WindowsUIButtonPanelActions.AutoSizeInLayoutControl = False
        WindowsUIButtonPanelActions.ForeColor = AbovoBlue
        WindowsUIButtonPanelActions.AppearanceButton.Normal.ForeColor = AbovoBlue
        WindowsUIButtonPanelActions.AppearanceButton.Normal.ForeColor = AbovoBlue

    End Sub
    Private Sub ProcessLinkElement()

        For Each Button In Me.WindowsUIButtonPanelActions.Buttons

            If TryCast(Button, DevExpress.XtraBars.Docking2010.WindowsUIButton) IsNot Nothing Then

                Dim ValidButton As DevExpress.XtraBars.Docking2010.WindowsUIButton = Button

                If ValidButton.Tag = "Return" Then

                    ValidButton.ToolTip = "Return to " & ActiveLinkElement.LinkReturnName
                    ValidButton.Visible = True

                    GoTo SectionSelect

                End If

            End If

        Next

SectionSelect:

        If ActiveLinkElement.LinkTargetSection IsNot Nothing Then

            For Each TabPage In XtraTabControlNewGIT.TabPages

                If Trim(TabPage.Text) = Trim(ActiveLinkElement.LinkTargetSection) Then

                    XtraTabControlNewGIT.SelectedTabPage = TabPage
                    Exit Sub

                End If

            Next

        End If

    End Sub
    Public Sub ClearLinks()

        For Each Button In Me.WindowsUIButtonPanelActions.Buttons
            If TryCast(Button, DevExpress.XtraBars.Docking2010.WindowsUIButton) IsNot Nothing Then

                Dim ValidButton As DevExpress.XtraBars.Docking2010.WindowsUIButton = Button

                If ValidButton.Tag = "Return" Then

                    ValidButton.ToolTip = ""
                    ValidButton.Visible = False
                    Exit Sub

                End If
            End If
        Next
    End Sub
    Public Sub AddLink(LinkTag As ElementInterfaceLinkTag)

        Me.ActiveLinkElement = LinkTag
        ProcessLinkElement()

    End Sub
    Private Sub WindowsUIButtonPanelSaveClose_ButtonClick(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelActions.ButtonClick

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)

        If ButSender Is Nothing Then

            Return

        End If

        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "History"

                ' Close the model and dispose of the interface
                HistoryManager.Visible = True
                HistoryManager.Show()
                HistoryManager.BringToFront()

            Case "Spreadsheet"

                ' Close the model and dispose of the interface
                ExcelModels(ModelID).ShowSpreadsheet(Me.ActiveSpreadsheet, Me)

            Case "MainMenu"

                ' Close the model and dispose of the interface
                If FormMainScreen.WindowState = FormWindowState.Minimized Then
                    FormMainScreen.WindowState = FormWindowState.Normal
                End If

                FormMainScreen.BringToFront()

            Case "Refresh"

                ' Close the model and dispose of the interface
                RefreshData()

            Case "Return"

                ' Open the parent interface, passing the return ID to allow it to navigate to the correct place
                Dim LinkTag As New ElementInterfaceLinkTag With {
                    .LinkType = "Link",
                    .LinkData = ActiveLinkElement.LinkReturnName,
                    .LinkGroupID = ActiveLinkElement.LinkReturnGroup,
                    .LinkReturnID = CSID,
                    .LinkReturnGroup = GSID,
                    .LinkTargetSection = Nothing,
                    .LinkReturnName = DITName
                }
                ExcelModels(ModelID).EventCoordinator.TriggerEvent("Link", LinkTag, ParentGroupForm)


        End Select

    End Sub

#End Region

#Region "Interface Handlers"
    Private Sub ComboOpen(sender, e)

        Dim CB As ComboBoxEdit = sender
        CB.ShowPopup()

    End Sub
    Sub GVPasting(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.ClipboardRowPastingEventArgs)

        Dim okays = e.GetValidValues
        Dim Errs = e.GetInvalidValues

    End Sub
    Private Sub DEHyperLinkLabel_linkClick(sender As Object, e As DevExpress.Utils.HyperlinkClickEventArgs)

        Dim LinkTag As ElementInterfaceLinkTag = TryCast(sender.Tag, ElementInterfaceLinkTag)

        If LinkTag Is Nothing Then

            MsgBox("Sorry, this link Is Not properly configured")
            Return

        End If

        Dim HyperlinkLabelControl As AbovoDEHyperlinkLabel = sender
        HyperlinkLabelControl.LinkVisited = True

        ExcelModels(ModelID).EventCoordinator.TriggerEvent("Link", LinkTag, ParentGroupForm)

    End Sub
    Private Sub Grid_ButtonClick(sender As Object, e As EventArgs)

        Dim GridTag As AttachedGridCommandButton =
            TryCast(sender.Tag, AttachedGridCommandButton)

        If GridTag Is Nothing Then

            MsgBox("Sorry, this button Is Not properly configured")
            Return

        End If

        Dim Trans As AbovoTransaction =
            ExcelModels(ModelID).EventCoordinator.TriggerEvent(
                "GridButton",
                GridTag,
                ParentGroupForm)

        If Trans Is Nothing OrElse Trans.BError Then Return

        Try

            'The specific structural routine holds the wait cursor while workbook
            'mutation is in progress. Reassert it here so it remains busy during
            'the subsequent interface rebuild as well.
            Me.UseWaitCursor = True
            Me.Cursor = Cursors.WaitCursor
            Cursor.Current = Cursors.WaitCursor

            RebuildAllSections()

        Finally

            Me.UseWaitCursor = False
            Me.Cursor = Cursors.Default
            Cursor.Current = Cursors.Default

        End Try

    End Sub
    Private Sub Interface_ButtonClick(sender As Object, e As EventArgs)

        Dim EventResult As AbovoTransaction =
            ExcelModels(ModelID).EventCoordinator.TriggerEvent(
                "Code",
                sender.Tag,
                ParentGroupForm)

        If EventResult Is Nothing OrElse EventResult.EventCancelled Then Return

        If EventResult.BError Then
            MessageBox.Show(
                Me,
                EventResult.StrResponseMessage,
                "Import Data",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            Return
        End If

        Dim ImportCommand As String = TryCast(sender.Tag, String)
        Dim IsDevelopmentImport As Boolean =
            String.Equals(ImportCommand, "ImportSingleDSA_File", StringComparison.Ordinal) OrElse
            String.Equals(ImportCommand, "ImportMultiDSA_Files", StringComparison.Ordinal) OrElse
            String.Equals(ImportCommand, "ImportConsolDSA_File", StringComparison.Ordinal) OrElse
            String.Equals(ImportCommand, "ImportDSA_Template", StringComparison.Ordinal)

        If String.Equals(ImportCommand, "ImportStockRentModel", StringComparison.Ordinal) OrElse
           String.Equals(ImportCommand, "ImportManagementServiceCosts", StringComparison.Ordinal) OrElse
           IsDevelopmentImport Then
            RebuildAllSections()
        End If

        If String.Equals(ImportCommand, "ImportManagementServiceCosts", StringComparison.Ordinal) OrElse
           IsDevelopmentImport Then
            MessageBox.Show(
                Me,
                EventResult.StrResponseMessage,
                "Import Data",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub Link_ButtonClick(sender As Object, e As EventArgs)

        Dim LinkTag As ElementInterfaceLinkTag = TryCast(sender.Tag, ElementInterfaceLinkTag)

        If LinkTag Is Nothing Then

            MsgBox("Sorry, this link Is Not properly configured")
            Return

        End If

        ExcelModels(ModelID).EventCoordinator.TriggerEvent("Link", sender.Tag, ParentGroupForm)

    End Sub


    Private Sub GridView_MouseWheel(sender As Object, e As MouseEventArgs)

        'AccordionControlM.Top += e.Delta
        DirectCast(e, DXMouseEventArgs).Handled = True

    End Sub
    Private Sub TooltipController_HyperlinkClick(sender As Object, e As DevExpress.Utils.HyperlinkClickEventArgs)

        Dim process As New Process()

        process.StartInfo.FileName = (e.Link)
        process.StartInfo.Verb = "open"
        process.StartInfo.WindowStyle = ProcessWindowStyle.Normal

        Try
            process.Start()
        Catch
            MsgBox("Sorry, the browser could Not be opened")
        End Try

    End Sub
    Sub VerifyDoubleClick(sender As Object, e As MouseEventArgs)



    End Sub
    Private Sub SetGridDisplayText(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs)

        'Dim ColTag As DataColumnTag = e.Column.Tag

        ''Select Case ColTag.DataType

        ''    Case "C"
        ''        Dim ciGB As CultureInfo = New CultureInfo("en-GB")
        ''        e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value))

        ''End Select

    End Sub

    Private Sub UndoLastAction()

        'ExcelModels(ModelID).UndoLastAction()
        RebuildAllSections()

    End Sub


    Private Sub KeydownListener(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode

            '  Ctrl + Shift + D
            Case Keys.P And (e.Shift And e.Control And Not e.Alt)
                MsgBox("Print")

                '  Ctrl + Z
            Case Keys.Z And (e.Control And Not e.Shift And Not e.Alt)

                MsgBox("Undo DIT")

                '  Shift and F1
            Case Keys.F1 And (e.Shift And Not e.Control And Not e.Alt)
                MsgBox("Help called")

            Case Keys.V And (e.Control And Not e.Shift And Not e.Alt)

                'GridControl/VGridControl Ctrl+V is handled by the dedicated
                'custom paste handlers attached when those controls are created.
                'Do not show the old diagnostic message here.
                Return

            Case Keys.C And (e.Control And Not e.Shift And Not e.Alt)

                MsgBox("Copy command")

        End Select

    End Sub

    Private Sub VGrid_Event_DoubleClick(ByVal sender As Object, ByVal e As EventArgs)

        DblClickCell = True

    End Sub

    '==========================================================================
    ' CUSTOM GRID / VGRID PASTE SUPPORT
    '==========================================================================
    ' XtraGrid has a native clipboard paste pipeline, but we deliberately
    ' intercept Ctrl+V at GridControl.ProcessGridKey so the spreadsheet and
    ' ChangeManager remain the authoritative write path.
    '
    ' VGridControl does not expose the same row-paste pipeline, so Ctrl+V is
    ' intercepted through KeyDown and routed into the same custom paste core.
    '
    ' Clipboard text from Excel and DevExpress grids is normally tab-delimited
    ' with CR/LF-delimited rows, which this parser accepts.

    Private Sub GridControl_ProcessGridKey(ByVal sender As Object, ByVal e As KeyEventArgs)

        If e.KeyCode <> Keys.V OrElse Not e.Control OrElse e.Alt Then Return

        Dim GC As GridControl = TryCast(sender, GridControl)
        If GC Is Nothing Then Return

        e.Handled = True
        e.SuppressKeyPress = True

        CustomPasteIntoDataGrid(GC)

    End Sub

    Private Sub VGrid_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)

        If e.KeyCode <> Keys.V OrElse Not e.Control OrElse e.Alt Then Return

        Dim VG As VGridControl = TryCast(sender, VGridControl)
        If VG Is Nothing Then Return

        e.Handled = True
        e.SuppressKeyPress = True

        CustomPasteIntoVGrid(VG)

    End Sub

    Private Function GetClipboardPasteMatrix() As List(Of String())

        Dim Result As New List(Of String())

        If Not Clipboard.ContainsText() Then Return Result

        Dim ClipboardText As String = Clipboard.GetText(TextDataFormat.UnicodeText)
        If String.IsNullOrEmpty(ClipboardText) Then Return Result

        Dim NormalisedText As String = ClipboardText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

        'Excel/DevExpress commonly leave one trailing newline on multi-cell copy.
        While NormalisedText.EndsWith(vbLf, StringComparison.Ordinal)
            NormalisedText = NormalisedText.Substring(0, NormalisedText.Length - 1)
        End While

        If NormalisedText.Length = 0 Then Return Result

        Dim ClipboardRows() As String = NormalisedText.Split(New String() {vbLf}, StringSplitOptions.None)

        For Each ClipboardRow As String In ClipboardRows
            Result.Add(ClipboardRow.Split(New Char() {ControlChars.Tab}, StringSplitOptions.None))
        Next

        Return Result

    End Function

    Private Function GetGridColumnIndex(ByVal Column As GridColumn) As Integer

        If Column Is Nothing Then Return -1

        Dim ColTag As DataColumnTag = TryCast(Column.Tag, DataColumnTag)
        If ColTag IsNot Nothing AndAlso Not String.IsNullOrEmpty(ColTag.ActiveColumnName) Then

            If ColTag.ActiveColumnName.StartsWith("Col_", StringComparison.OrdinalIgnoreCase) Then
                Dim Result As Integer
                If Integer.TryParse(ColTag.ActiveColumnName.Substring(4), Result) Then Return Result
            End If

        End If

        If Not String.IsNullOrEmpty(Column.FieldName) AndAlso
           Column.FieldName.StartsWith("Col_", StringComparison.OrdinalIgnoreCase) Then

            Dim Result As Integer
            If Integer.TryParse(Column.FieldName.Substring(4), Result) Then Return Result

        End If

        Return -1

    End Function

    Private Function ConvertPastedTextValue(ByVal RawValue As String, ByVal ColTag As DataColumnTag, ByRef ConvertedValue As Object) As Boolean

        If ColTag Is Nothing Then Return False

        'Blank clipboard cells are intentionally represented as Nothing so the
        'existing PushDSData null path clears the underlying DataCellPoint.

        If RawValue Is Nothing OrElse RawValue.Length = 0 Then
            ConvertedValue = Nothing
            Return True
        End If

        Select Case ColTag.DataType

            Case "S"
                ConvertedValue = RawValue
                Return True

            Case "B"
                Dim BoolValue As Boolean
                If Boolean.TryParse(RawValue, BoolValue) Then
                    ConvertedValue = If(BoolValue, 1, 0)
                    Return True
                End If

                Dim BoolNumber As Double
                If Double.TryParse(RawValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, BoolNumber) Then
                    ConvertedValue = If(BoolNumber = 0, 0, 1)
                    Return True
                End If

                Return False

            Case "I", "Y"
                Dim IntValue As Integer
                If Integer.TryParse(RawValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, IntValue) Then
                    ConvertedValue = IntValue
                    Return True
                End If
                Return False

            Case "N", "C", "M", "SM", "R", "P", "D"
                Dim NumericText As String = RawValue.Trim()
                Dim IsPercent As Boolean = NumericText.EndsWith("%", StringComparison.Ordinal)

                If IsPercent Then NumericText = NumericText.Substring(0, NumericText.Length - 1).Trim()

                Dim DoubleValue As Double
                If Double.TryParse(NumericText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, DoubleValue) Then
                    If IsPercent Then DoubleValue /= 100.0
                    ConvertedValue = DoubleValue
                    Return True
                End If

                Return False

            Case Else
                ConvertedValue = RawValue
                Return True

        End Select

    End Function

    Private Function CanPasteToDataPoint(ByVal DataSet As DataCellRange,
                                         ByVal DataRowIndex As Integer,
                                         ByVal DataColumnIndex As Integer) As Boolean

        If DataSet Is Nothing Then Return False
        If DataRowIndex < 0 OrElse DataRowIndex >= DataSet.DataRows.Count Then Return False
        If DataColumnIndex < 0 OrElse DataColumnIndex >= DataSet.DataColumns.Count Then Return False

        Dim ColTag As DataColumnTag = DataSet.DataColumns(DataColumnIndex).ColumnTag
        If ColTag Is Nothing Then Return False

        If DataSet.RO OrElse ColTag.IsReadOnly OrElse ColTag.IsCalculated OrElse ColTag.IsDummyColumn Then Return False
        If DataSet.DataRows(DataRowIndex).IsSpacerRow OrElse DataSet.DataRows(DataRowIndex).IsControlRow Then Return False

        Dim DP As CellDataPoint = DataSet.DataRows(DataRowIndex).DataCells(DataColumnIndex)
        If DP Is Nothing Then Return False
        If DP.IsLocked Then Return False

        If ColTag.HasRules Then
            If ActiveSpreadsheet IsNot Nothing AndAlso
               ActiveSpreadsheet.Range(DP.SourceAddress).Fill.PatternType <> PatternType.Solid Then Return False
        End If

        Return True

    End Function

    Private Sub CustomPasteIntoDataGrid(ByVal Grid As GridControl)

        Dim View As GridView = TryCast(Grid.FocusedView, GridView)
        If View Is Nothing OrElse View.FocusedColumn Is Nothing Then Return

        Dim UBS As AbovoUnboundSource = TryCast(Grid.DataSource, AbovoUnboundSource)
        If UBS Is Nothing Then Return

        Dim DSIndex As Integer = UBS.UBSTag.DSIndex
        If DSIndex < 0 OrElse DSIndex >= DataPres.DataSets.Count Then Return

        Dim TargetDataSet As DataCellRange = DataPres.DataSets(DSIndex)
        If TargetDataSet.RO Then Return

        Dim PasteMatrix As List(Of String()) = GetClipboardPasteMatrix()
        If PasteMatrix.Count = 0 Then Return

        'Important: use the stable data-source index rather than FocusedRowHandle,
        'because the view may be sorted/filtered.
        Dim StartDataRow As Integer = View.GetFocusedDataSourceRowIndex()
        Dim StartDataColumn As Integer = GetGridColumnIndex(View.FocusedColumn)

        If StartDataRow < 0 OrElse StartDataColumn < 0 Then Return

        ApplyPasteMatrix(DSIndex, TargetDataSet, PasteMatrix, StartDataRow, StartDataColumn, False)

    End Sub

    Private Sub CustomPasteIntoVGrid(ByVal VertGrid As VGridControl)

        If VertGrid.FocusedRow Is Nothing Then Return

        Dim UBS As AbovoUnboundSource = TryCast(VertGrid.DataSource, AbovoUnboundSource)
        If UBS Is Nothing Then Return

        Dim DSIndex As Integer = UBS.UBSTag.DSIndex
        If DSIndex < 0 OrElse DSIndex >= DataPres.DataSets.Count Then Return

        Dim TargetDataSet As DataCellRange = DataPres.DataSets(DSIndex)
        If TargetDataSet.RO Then Return

        Dim PasteMatrix As List(Of String()) = GetClipboardPasteMatrix()
        If PasteMatrix.Count = 0 Then Return

        Dim StartDataRow As Integer = VertGrid.FocusedRecord
        Dim StartDataColumn As Integer = GetVGridColumnIndex(VertGrid.FocusedRow)

        If StartDataRow < 0 OrElse StartDataColumn < 0 Then Return

        'VGrid is visually transposed relative to XtraGrid:
        '  clipboard rows    -> successive VGrid editor rows (dataset columns)
        '  clipboard columns -> successive VGrid records (dataset rows)
        ApplyPasteMatrix(DSIndex, TargetDataSet, PasteMatrix, StartDataRow, StartDataColumn, True)

    End Sub

    Private Sub ApplyPasteMatrix(ByVal DSIndex As Integer,
                                 ByVal TargetDataSet As DataCellRange,
                                 ByVal PasteMatrix As List(Of String()),
                                 ByVal StartDataRow As Integer,
                                 ByVal StartDataColumn As Integer,
                                 ByVal TransposeForVGrid As Boolean)

        If TargetDataSet Is Nothing OrElse PasteMatrix Is Nothing OrElse PasteMatrix.Count = 0 Then Return

        Dim AnyChanged As Boolean = False

        Me.Cursor = Cursors.WaitCursor

        Try

            For ClipboardRowIndex As Integer = 0 To PasteMatrix.Count - 1

                Dim ClipboardRow() As String = PasteMatrix(ClipboardRowIndex)

                For ClipboardColumnIndex As Integer = 0 To ClipboardRow.Length - 1

                    Dim TargetRowIndex As Integer
                    Dim TargetColumnIndex As Integer

                    If TransposeForVGrid Then
                        TargetRowIndex = StartDataRow + ClipboardColumnIndex
                        TargetColumnIndex = StartDataColumn + ClipboardRowIndex
                    Else
                        TargetRowIndex = StartDataRow + ClipboardRowIndex
                        TargetColumnIndex = StartDataColumn + ClipboardColumnIndex
                    End If

                    If TargetRowIndex >= TargetDataSet.DataRows.Count OrElse
                       TargetColumnIndex >= TargetDataSet.DataColumns.Count Then Continue For

                    If Not CanPasteToDataPoint(TargetDataSet, TargetRowIndex, TargetColumnIndex) Then Continue For

                    Dim ColTag As DataColumnTag = TargetDataSet.DataColumns(TargetColumnIndex).ColumnTag
                    Dim ConvertedValue As Object = Nothing

                    If Not ConvertPastedTextValue(ClipboardRow(ClipboardColumnIndex), ColTag, ConvertedValue) Then
                        'For now invalid clipboard values are skipped. This is a deliberate
                        'placeholder for a future aggregated paste-validation report.
                        Continue For
                    End If

                    Dim SourceDataPoint As CellDataPoint =
                        TargetDataSet.DataRows(TargetRowIndex).DataCells(TargetColumnIndex)

                    'Expose the resolved spreadsheet target to a dedicated paste hook.
                    'This is where paste-specific mapping/validation can be added later.
                    CustomPasteCellHook(DSIndex,
                                        TargetRowIndex,
                                        TargetColumnIndex,
                                        SourceDataPoint.SourceSheet,
                                        SourceDataPoint.SourceAddress,
                                        ClipboardRow(ClipboardColumnIndex),
                                        ConvertedValue)

                    'This uses the same DataChangeEvent / ChangeManager path as normal
                    'single-cell edits, but suppresses a full UI refresh for every cell.
                    PushDSData(DSIndex, TargetRowIndex, TargetColumnIndex, ConvertedValue, False)
                    AnyChanged = True

                Next

            Next

            If AnyChanged Then
                CustomPastePostProcess(DSIndex, StartDataRow, StartDataColumn)
                UpdateRules(DSIndex)
                UpdateCalcs(DSIndex)
                RefreshData()
            End If

        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    'Per-cell extension point. SourceSheet/SourceAddress are already resolved
    'from the underlying CellDataPoint before ChangeManager is called.
    Private Sub CustomPasteCellHook(ByVal DSIndex As Integer,
                                    ByVal DataRowIndex As Integer,
                                    ByVal DataColumnIndex As Integer,
                                    ByVal SourceSheet As String,
                                    ByVal SourceAddress As String,
                                    ByVal RawClipboardValue As String,
                                    ByVal ConvertedValue As Object)

        'Intentionally empty for now.
        'Examples of future use:
        '  - special paste validation
        '  - logging/preview of target worksheet + cell address
        '  - formula-aware or named-range-aware paste rules
        '  - source-specific handling for Excel vs another Abovo/DevExpress grid

    End Sub

    'Future batch-level extension point for paste-specific transformations, validation,
    'cross-grid metadata, or special spreadsheet addressing rules.
    Private Sub CustomPastePostProcess(ByVal DSIndex As Integer,
                                       ByVal FirstDataRow As Integer,
                                       ByVal FirstDataColumn As Integer)

        'Intentionally empty for now.

    End Sub

    Private Function IsJointVenturePairedBand(ByVal BandID As String) As Boolean

        If String.IsNullOrWhiteSpace(BandID) Then Return False

        Select Case BandID.Trim()

            Case "Investments",
                 "Share of Surplus/Deficit",
                 "JV Cash Investment Repayments",
                 "JV Interest Rate"

                Return True

        End Select

        Return False

    End Function

    Private Function GetJointVentureRange(ByVal RangeName As String) As DevExpress.Spreadsheet.CellRange
        Dim DefinedName = ActiveWorkbook.DefinedNames.GetDefinedName(RangeName)
        If DefinedName Is Nothing Then
            Throw New InvalidOperationException("Joint Venture named range '" & RangeName & "' was not found.")
        End If
        Return DefinedName.Range
    End Function

    Private Function GetJointVentureVectorCell(ByVal SourceRange As DevExpress.Spreadsheet.CellRange,
                                               ByVal Index As Integer) As DevExpress.Spreadsheet.Cell
        If SourceRange.RowCount = 1 Then Return SourceRange(0, Index)
        If SourceRange.ColumnCount = 1 Then Return SourceRange(Index, 0)
        'JV master/opening ranges are expected to be vectors.  If a legacy model
        'contains a wider range, use the first row as the record vector.
        Return SourceRange(0, Index)
    End Function

    Private Function GetJointVentureCellValue(ByVal Cell As DevExpress.Spreadsheet.Cell,
                                              ByVal DataFormat As String) As Object
        If Cell Is Nothing OrElse Cell.Value.IsEmpty Then Return DBNull.Value

        Select Case DataFormat
            Case "I", "Y"
                Return CInt(Cell.Value.NumericValue)
            Case "N", "P", "C", "M", "SM", "R"
                Return CDbl(Cell.Value.NumericValue)
            Case Else
                Return Cell.DisplayText
        End Select
    End Function

    Private Sub SetJointVentureBinding(ByVal Grid As VGridControl,
                                       ByVal FieldName As String,
                                       ByVal RecordIndex As Integer,
                                       ByVal Cell As DevExpress.Spreadsheet.Cell,
                                       ByVal DataFormat As String,
                                       Optional ByVal IsReadOnly As Boolean = False,
                                       Optional ByVal HasRules As Boolean = False)
        If Cell Is Nothing Then Return

        Dim Map As Dictionary(Of String, JointVentureCellBinding) = Nothing
        If Not JointVentureBindings.TryGetValue(Grid, Map) Then
            Map = New Dictionary(Of String, JointVentureCellBinding)(StringComparer.Ordinal)
            JointVentureBindings(Grid) = Map
        End If

        Map(FieldName & "|" & RecordIndex.ToString) =
            New JointVentureCellBinding With {
                .SourceSheet = Cell.Worksheet.Name,
                .SourceAddress = Cell.GetReferenceA1,
                .DataFormat = DataFormat,
                .IsReadOnly = IsReadOnly,
                .HasRules = HasRules
            }
    End Sub

    Private Function IsJointVentureRecordDefined(ByVal RecordIndex As Integer) As Boolean

        'Records 0 and 1 are the synthetic Description and Year records.
        'Real Joint Venture records begin at index 2.
        If RecordIndex < 2 Then Return True

        Dim JVNames As DevExpress.Spreadsheet.CellRange =
            GetJointVentureRange("IC_JointVenture_01")

        Dim JVIndex As Integer = RecordIndex - 2

        Dim JVCount As Integer =
            If(JVNames.RowCount = 1,
               JVNames.ColumnCount,
               JVNames.RowCount)

        If JVIndex < 0 OrElse JVIndex >= JVCount Then Return False

        Dim NameCell As DevExpress.Spreadsheet.Cell =
            GetJointVentureVectorCell(
                JVNames,
                JVIndex)

        If NameCell Is Nothing OrElse NameCell.Value.IsEmpty Then Return False

        Return Not String.IsNullOrWhiteSpace(NameCell.DisplayText)

    End Function

    Private Function IsJointVentureRuleCellLocked(ByVal Binding As JointVentureCellBinding) As Boolean

        If Binding Is Nothing OrElse Not Binding.HasRules Then Return False

        Dim SourceCell As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets(Binding.SourceSheet).
                Cells(Binding.SourceAddress)

        If SourceCell Is Nothing Then Return True

        'This deliberately mirrors DataManager's Funding/normal-grid rule:
        '
        '   HasRule=True + Solid conditional-format fill = editable
        '   HasRule=True + non-Solid fill             = locked
        '
        'The workbook remains the authority for whether a rule-controlled cell
        'is currently active.
        Return SourceCell.Fill.PatternType <> PatternType.Solid

    End Function

    Private Function IsJointVentureCellLocked(ByVal Grid As VGridControl,
                                              ByVal Row As BaseRow,
                                              ByVal RecordIndex As Integer,
                                              ByVal Binding As JointVentureCellBinding) As Boolean

        If Binding Is Nothing Then Return True
        If Binding.IsReadOnly Then Return True

        Dim RowTag As JointVentureVGridRowTag =
            TryCast(Row.Tag, JointVentureVGridRowTag)

        'JV Name itself must remain editable so an unused JV column can be
        'activated by entering its name. Every lower cell in that JV record is
        'inactive until the master name exists.
        If RecordIndex >= 2 AndAlso
           (RowTag Is Nothing OrElse
            Not String.Equals(RowTag.RowKind,
                              "JVName",
                              StringComparison.OrdinalIgnoreCase)) Then

            If Not IsJointVentureRecordDefined(RecordIndex) Then Return True

        End If

        If IsJointVentureRuleCellLocked(Binding) Then Return True

        Return False

    End Function

    Private Function GetJointVentureBinding(ByVal Grid As VGridControl,
                                            ByVal Row As BaseRow,
                                            ByVal RecordIndex As Integer) As JointVentureCellBinding

        Dim EditorRow As EditorRow = TryCast(Row, EditorRow)
        If EditorRow Is Nothing Then Return Nothing

        Dim Map As Dictionary(Of String, JointVentureCellBinding) = Nothing
        If Not JointVentureBindings.TryGetValue(Grid, Map) Then Return Nothing

        Dim Key As String =
            EditorRow.Properties.FieldName &
            "|" &
            RecordIndex.ToString

        Dim Binding As JointVentureCellBinding = Nothing

        If Not Map.TryGetValue(Key, Binding) Then Return Nothing

        Return Binding

    End Function

    Private Function CreateJointVentureAddColumnsPanel(ByVal Grid As VGridControl,
                                                       ByVal SetSectionID As Integer) As System.Windows.Forms.Panel

        Dim ButtonPanel As New System.Windows.Forms.Panel With {
            .AutoSize = False,
            .Dock = System.Windows.Forms.DockStyle.None,
            .Anchor = System.Windows.Forms.AnchorStyles.Top Or
                      System.Windows.Forms.AnchorStyles.Left Or
                      System.Windows.Forms.AnchorStyles.Right,
            .Tag = Grid,
            .Margin = New System.Windows.Forms.Padding(0),
            .Padding = New System.Windows.Forms.Padding(
                CInt(DefaultTablePanelPadding.Left),
                CInt(DefaultTablePanelPadding.Top),
                CInt(DefaultTablePanelPadding.Right),
                CInt(DefaultTablePanelPadding.Bottom))
        }

        Dim CommandButton As New DevExpress.XtraEditors.SimpleButton With {
            .AllowHtmlDraw = DefaultBoolean.True,
            .Text = "Add Joint Venture Columns",
            .ToolTip = "Add joint venture columns"
        }

        CommandButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap

        Dim AddJVAction As New ActionToken With {
            .ActionType = "NRCI",
            .ActionStrData1 = "IC_JointVenture_01",
            .ActionNR = "IC_JointVenture_01",
            .ActionNumber1 = SetSectionID,
            .ActionDescription = "Add Joint Venture Columns"
        }

        CommandButton.Tag = AddJVAction

        Dim BestSize As Size = CommandButton.CalcBestSize()
        CommandButton.Width = Math.Max(BestSize.Width, 170)
        CommandButton.Height = BestSize.Height

        'Match the Funding VGrid command strip: one command positioned at the
        'right-hand edge of the actual VGrid content area.
        CommandButton.Left = Math.Max(
            CInt(DefaultTablePanelPadding.Left),
            Grid.Width -
            CommandButton.Width -
            CInt(DefaultTablePanelPadding.Right))

        CommandButton.Top = CInt(DefaultTablePanelPadding.Top)

        AddHandler CommandButton.Click,
            AddressOf JointVenture_AddColumnsButtonClick

        ButtonPanel.Controls.Add(CommandButton)

        ButtonPanel.Height =
            CommandButton.Height +
            DefaultTablePanelPadding.Top +
            DefaultTablePanelPadding.Bottom

        Return ButtonPanel

    End Function

    Private Sub JointVenture_AddColumnsButtonClick(ByVal sender As Object,
                                                   ByVal e As EventArgs)

        Dim Button As DevExpress.XtraEditors.SimpleButton =
            TryCast(sender, DevExpress.XtraEditors.SimpleButton)

        If Button Is Nothing Then Return

        Dim Act As ActionToken =
            TryCast(Button.Tag, ActionToken)

        If Act Is Nothing Then Return

        RunAction(Act)

    End Sub

    Private Function BuildJointVentureExcelStyleVGrid(ByVal SetSectionID As Integer) As VGridControl

        Dim JVNames As DevExpress.Spreadsheet.CellRange = GetJointVentureRange("IC_JointVenture_01")
        Dim OpeningBalances As DevExpress.Spreadsheet.CellRange = GetJointVentureRange("Rep_JointVenture_010")

        Dim JVCount As Integer = If(JVNames.RowCount = 1, JVNames.ColumnCount, JVNames.RowCount)
        Dim RecordCount As Integer = JVCount + 2 'Description + Year + JV records

        Dim Table As New DataTable("JointVentureExcelStyle")
        For RecordIndex As Integer = 0 To RecordCount - 1
            Table.Rows.Add(Table.NewRow())
        Next

        VertGridCount += 1
        ReDim Preserve VertGridControls(VertGridCount)

        Dim Grid As New VGridControl() With {
            .Name = "VGridControl_JointVenture_" & VertGridCount.ToString,
            .Parent = Me,
            .Dock = DockStyle.None,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
        }
        VertGridControls(VertGridCount) = Grid
        CurrentAbovoTabPage.AddVGrid(Grid)
        JointVentureTables(Grid) = Table
        JointVentureBindings(Grid) = New Dictionary(Of String, JointVentureCellBinding)(StringComparer.Ordinal)

        'Repository items are created once and merely selected per cell in
        'CustomRecordCellEdit.  This is the DevExpress-supported way to give
        'Description, Year and JV value records different editors in one VGrid row.
        Dim TextEdit As New RepositoryItemTextEdit()
        Dim YearEdit As RepositoryItemComboBox =
            RepositaryItems.GetEditor("Rep_OrdinalYears", ModelID).RetCombo
        Dim MoneyEdit As New RepositoryItemTextEdit()
        MoneyEdit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
        MoneyEdit.Mask.EditMask = "c0"
        MoneyEdit.UseMaskAsDisplayFormat = True
        Dim PercentEdit As New RepositoryItemSpinEdit()
        PercentEdit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
        PercentEdit.Mask.EditMask = "p2"
        PercentEdit.DisplayFormat.FormatType = FormatType.Numeric
        PercentEdit.DisplayFormat.FormatString = "p2"
        PercentEdit.UseMaskAsDisplayFormat = True

        Grid.RepositoryItems.Add(TextEdit)
        Grid.RepositoryItems.Add(YearEdit)
        Grid.RepositoryItems.Add(MoneyEdit)
        Grid.RepositoryItems.Add(PercentEdit)
        JointVentureTextEditors(Grid) = TextEdit
        JointVentureYearEditors(Grid) = YearEdit
        JointVentureMoneyEditors(Grid) = MoneyEdit
        JointVenturePercentEditors(Grid) = PercentEdit

        Dim FieldCounter As Integer = -1
        Dim CategoryRows As New List(Of CategoryRow)

        Dim AddField =
            Function(CategoryName As String,
                     Caption As String,
                     RowKind As String,
                     ValueFormat As String,
                     DescriptionCell As DevExpress.Spreadsheet.Cell,
                     YearCell As DevExpress.Spreadsheet.Cell,
                     ValueRange As DevExpress.Spreadsheet.CellRange,
                     ValueRowIndex As Integer,
                     DescriptionHasRules As Boolean,
                     YearHasRules As Boolean,
                     ValueHasRules As Boolean) As String

                FieldCounter += 1
                Dim FieldName As String = "JVField_" & FieldCounter.ToString
                Table.Columns.Add(FieldName, GetType(Object))

                If DescriptionCell IsNot Nothing Then
                    Table.Rows(0)(FieldName) = GetJointVentureCellValue(DescriptionCell, "S")
                    SetJointVentureBinding(Grid, FieldName, 0, DescriptionCell, "S",
                                             False, DescriptionHasRules)
                End If
                If YearCell IsNot Nothing Then
                    Table.Rows(1)(FieldName) = GetJointVentureCellValue(YearCell, "I")
                    SetJointVentureBinding(Grid, FieldName, 1, YearCell, "I",
                                             False, YearHasRules)
                End If

                If ValueRange IsNot Nothing Then
                    For JVIndex As Integer = 0 To JVCount - 1
                        If ValueRowIndex < ValueRange.RowCount AndAlso JVIndex < ValueRange.ColumnCount Then
                            Dim ValueCell As DevExpress.Spreadsheet.Cell = ValueRange(ValueRowIndex, JVIndex)
                            Table.Rows(JVIndex + 2)(FieldName) = GetJointVentureCellValue(ValueCell, ValueFormat)
                            SetJointVentureBinding(Grid, FieldName, JVIndex + 2, ValueCell, ValueFormat,
                                                     False, ValueHasRules)
                        End If
                    Next
                End If

                Dim VRow As New EditorRow(FieldName) With {
                    .Height = IdealGridRowHeight,
                    .Tag = New JointVentureVGridRowTag With {
                        .CategoryName = CategoryName,
                        .RowKind = RowKind,
                        .ValueDataFormat = ValueFormat
                    }
                }
                VRow.Properties.Caption = Caption
                VRow.Properties.FieldName = FieldName
                Grid.Rows.Add(VRow)

                Return FieldName
            End Function

        'Top two worksheet rows.  The first two metadata records intentionally
        'remain blank, exactly as they are on the Excel sheet.
        'JV Name is the master switch for each real JV record and is therefore
        'not itself rule-controlled. Opening Balance is <HasRule>TRUE</HasRule>
        'in Structure.xml and follows the workbook's conditional formatting.
        AddField("Joint Ventures", "JV Name", "JVName", "S",
                 Nothing, Nothing, JVNames, 0,
                 False, False, False)

        AddField("Joint Ventures", "Opening Balance £'000", "OpeningBalance", "C",
                 Nothing, Nothing, OpeningBalances, 0,
                 False, False, True)

        Dim AddEventBand =
            Sub(CategoryName As String,
                HeaderRangeName As String,
                ValueRangeName As String,
                ValueFormat As String,
                InterestMode As Boolean)

                Dim HeaderRange As DevExpress.Spreadsheet.CellRange = GetJointVentureRange(HeaderRangeName)
                Dim ValueRange As DevExpress.Spreadsheet.CellRange = GetJointVentureRange(ValueRangeName)
                Dim EntryCount As Integer = Math.Min(HeaderRange.RowCount, ValueRange.RowCount)

                For EntryIndex As Integer = 0 To EntryCount - 1
                    Dim DescriptionCell As DevExpress.Spreadsheet.Cell = Nothing
                    Dim YearCell As DevExpress.Spreadsheet.Cell = Nothing

                    If InterestMode Then
                        If HeaderRange.ColumnCount > 0 Then YearCell = HeaderRange(EntryIndex, 0)
                    Else
                        If HeaderRange.ColumnCount > 0 Then DescriptionCell = HeaderRange(EntryIndex, 0)
                        If HeaderRange.ColumnCount > 1 Then YearCell = HeaderRange(EntryIndex, 1)
                    End If

                    Dim Caption As String = " "

                    If InterestMode AndAlso EntryIndex = 0 Then
                        Caption = "Interest Rate"
                    End If

                    Dim CreatedFieldName As String =
                        AddField(CategoryName,
                                 Caption,
                                 If(InterestMode, "Interest", "Event"),
                                 ValueFormat,
                                 DescriptionCell,
                                 YearCell,
                                 ValueRange,
                                 EntryIndex,
                                 False,  'Description
                                 True,   'Year / Interest From Year
                                 True)   'Per-JV values

                    'The interest-rate header range contains the actual "from
                    'year" value, but not a separate Description field. Record 0
                    'is our synthetic descriptor column, so show a presentation-
                    'only label there using the EXACT field name AddField created.
                    'There is deliberately no JointVentureCellBinding for this
                    'synthetic value, so it can never write back to the workbook.
                    If InterestMode AndAlso
                       Not String.IsNullOrEmpty(CreatedFieldName) Then

                        Table.Rows(0)(CreatedFieldName) = "From year"

                    End If

                Next
            End Sub

        AddEventBand("Investments", "IR_JointVenture_01a", "Rep_JointVenture_030", "M", False)
        AddEventBand("Share of Surplus/Deficit", "IR_JointVenture_02a", "Rep_JointVenture_050", "M", False)
        AddEventBand("JV Cash Investment Repayments", "IR_JointVenture_03a", "Rep_JointVenture_070", "M", False)
        AddEventBand("JV Interest Rate", "IR_JointVenture_04a", "Rep_JointVenture_080", "P", True)

        Grid.DataSource = Table
        Grid.ForceInitialize()

        'Create worksheet-like categories.
        '
        'The Joint Ventures category is only a visual grouping for JV Name and
        'Opening Balance. Adding another JV is a physical WORKBOOK COLUMN action
        'and therefore belongs to the whole VGrid, not to this category.
        '
        'The four event categories genuinely add worksheet rows, so they retain
        'their NRRI category actions.
        Dim CategoryDefinitions As New List(Of Tuple(Of String, String, String)) From {
            Tuple.Create("Joint Ventures", "", ""),
            Tuple.Create("Investments", "NRRI", "IR_JointVenture_01a"),
            Tuple.Create("Share of Surplus/Deficit", "NRRI", "IR_JointVenture_02a"),
            Tuple.Create("JV Cash Investment Repayments", "NRRI", "IR_JointVenture_03a"),
            Tuple.Create("JV Interest Rate", "NRRI", "IR_JointVenture_04a")
        }

        For Each Definition In CategoryDefinitions
            Dim Cat As New CategoryRow("JVCategory_" & CategoryRows.Count.ToString) With {.Height = 40}
            Cat.Properties.Caption = Definition.Item1
            Dim HasCategoryAction As Boolean =
                Not String.IsNullOrWhiteSpace(Definition.Item2)

            Dim Act As ActionToken = Nothing

            If HasCategoryAction Then
                Act = New ActionToken With {
                    .ActionType = Definition.Item2,
                    .ActionStrData1 = Definition.Item3,
                    .ActionNR = Definition.Item3,
                    .ActionNumber1 = SetSectionID,
                    .ActionDescription = "Add Lines"
                }
            End If

            'JV categories use the VGrid's normal neutral separators.
            'The thick wheat border came from the banded-grid visual language
            'and is unnecessary here.
            Cat.Tag = New BandTag With {
                .ID = VertGridCount,
                .HasActions = HasCategoryAction,
                .ActionNR = Definition.Item3,
                .ActionDescription = If(HasCategoryAction, "Add Lines", Nothing),
                .ActionToken = Act,
                .HighLightColour = Color.WhiteSmoke,
                .DoBorder = False,
                .ButtonObjectState = ObjectState.Normal
            }
            Grid.Rows.Add(Cat)
            CategoryRows.Add(Cat)

            Dim RowsToMove As New List(Of BaseRow)
            For Each Row As BaseRow In Grid.Rows
                Dim JRTag As JointVentureVGridRowTag = TryCast(Row.Tag, JointVentureVGridRowTag)
                If JRTag IsNot Nothing AndAlso String.Equals(JRTag.CategoryName, Definition.Item1, StringComparison.Ordinal) Then
                    RowsToMove.Add(Row)
                End If
            Next
            For Each Row As BaseRow In RowsToMove
                Grid.MoveRow(Row, Cat, False)
            Next
        Next

        For Each Cat In CategoryRows
            Dim CatTag As BandTag = TryCast(Cat.Tag, BandTag)

            'Joint Ventures has no category action.  Its structural action is the
            'whole-grid Add Joint Venture Columns command above the VGrid.
            If CatTag Is Nothing OrElse
               Not CatTag.HasActions OrElse
               CatTag.ActionToken Is Nothing Then

                Continue For

            End If

            Dim Ext As New VGridCategoryButtonExtender(Grid, Cat, Me, "", SetSectionID,
                                                        Cat.Properties.Caption, CatTag.ActionToken)
            Ext.AddCustomButton()
            VGridCategoryExtenders.Add(Ext)

            Dim SectionExtenders As List(Of VGridCategoryButtonExtender) = Nothing
            If Not VGridCategoryExtendersBySection.TryGetValue(SetSectionID, SectionExtenders) Then
                SectionExtenders = New List(Of VGridCategoryButtonExtender)
                VGridCategoryExtendersBySection(SetSectionID) = SectionExtenders
            End If

            SectionExtenders.Add(Ext)
        Next

        Formatter.FormatVertGrid(Grid)
        Grid.RecordWidth = 145
        Grid.RowHeaderWidth = 220
        Grid.OptionsBehavior.Editable = True
        Grid.ScrollVisibility = DevExpress.XtraVerticalGrid.ScrollVisibility.Horizontal

        AddHandler Grid.CustomRecordCellEdit, AddressOf JointVenture_CustomRecordCellEdit
        AddHandler Grid.CustomRecordCellEditForEditing, AddressOf JointVenture_CustomRecordCellEdit
        AddHandler Grid.CustomDrawRowValueCell, AddressOf JointVenture_CustomDrawCell
        AddHandler Grid.Paint, AddressOf JointVenture_PostPaint
        AddHandler Grid.Disposed, AddressOf JointVenture_VGridDisposed
        AddHandler Grid.ShowingEditor, AddressOf JointVenture_ShowingEditor
        AddHandler Grid.CellValueChanged, AddressOf JointVenture_CellValueChanged
        AddHandler Grid.KeyDown, AddressOf VGrid_KeyDown
        AddHandler Grid.EditorKeyDown, AddressOf VGrid_KeyDown

        Grid.BestFit()
        Return Grid

    End Function

    Private Sub JointVenture_VGridDisposed(ByVal sender As Object,
                                              ByVal e As EventArgs)

        Dim Grid As VGridControl =
            TryCast(sender, VGridControl)

        If Grid Is Nothing Then Return

        JointVentureSeparatorSuppressions.Remove(Grid)

    End Sub

    Private Sub JointVenture_CustomRecordCellEdit(ByVal sender As Object,
                                                   ByVal e As DevExpress.XtraVerticalGrid.Events.GetCustomRowCellEditEventArgs)
        Dim Grid As VGridControl = TryCast(sender, VGridControl)
        If Grid Is Nothing Then Return
        Dim Tag As JointVentureVGridRowTag = TryCast(e.Row.Tag, JointVentureVGridRowTag)
        If Tag Is Nothing Then Return

        If e.RecordIndex = 0 Then
            If Tag.RowKind = "Event" Then e.RepositoryItem = JointVentureTextEditors(Grid)
            Return
        End If

        If e.RecordIndex = 1 Then
            If Tag.RowKind = "Event" OrElse Tag.RowKind = "Interest" Then
                e.RepositoryItem = JointVentureYearEditors(Grid)
            End If
            Return
        End If

        Select Case Tag.ValueDataFormat
            Case "P"
                e.RepositoryItem = JointVenturePercentEditors(Grid)
            Case "C", "M", "SM", "N", "R"
                e.RepositoryItem = JointVentureMoneyEditors(Grid)
            Case Else
                e.RepositoryItem = JointVentureTextEditors(Grid)
        End Select
    End Sub

    Private Sub JointVenture_CustomDrawCell(
        ByVal sender As Object,
        ByVal e As DevExpress.XtraVerticalGrid.Events.CustomDrawRowValueCellEventArgs)

        Dim Grid As VGridControl = TryCast(sender, VGridControl)

        If Grid Is Nothing OrElse e.Row Is Nothing Then Return

        '--------------------------------------------------------------
        ' Synthetic category footer
        '
        'VGridCategoryButtonExtender creates one unbound EditorRow as the
        'category footer.  Treat its record area as ONE continuous footer
        'strip rather than a sequence of data cells. DevExpress paints the
        'record separators after the individual cell custom draw, so queue the
        'vertical separator segments for removal in the final Paint pass.
        'The footer extender continues to own the button/text in the row-header
        'area; this handler only cleans up the record/value area.
        '--------------------------------------------------------------
        Dim FooterTag As VGridCategoryFooterTag =
            TryCast(e.Row.Tag, VGridCategoryFooterTag)

        If FooterTag IsNot Nothing Then

            Dim FooterBackColor As Color = Grid.Appearance.Category.BackColor

            If FooterBackColor = Color.Empty Then
                FooterBackColor = Color.WhiteSmoke
            End If

            e.Cache.FillRectangle(FooterBackColor, e.Bounds)

            'The separator itself is painted by the VGrid outside this cell's
            'custom-draw operation. Queue the LEFT separator of every record cell
            'for a final-pass erase. This removes the boundary between the row
            'header and first record as well as all internal record separators,
            'while retaining the grid's outer-right edge.
            QueueJointVentureSeparatorSuppression(
                Grid,
                "Footer|" &
                e.Row.Name &
                "|" &
                e.RecordIndex.ToString &
                "|L",
                e.Bounds.Left,
                e.Bounds.Top + 1,
                e.Bounds.Bottom - 1,
                FooterBackColor)

            e.Handled = True
            Return

        End If

        Dim JVRowTag As JointVentureVGridRowTag =
            TryCast(e.Row.Tag, JointVentureVGridRowTag)

        '--------------------------------------------------------------
        ' Blank Description / Year area beside JV Name and Opening Balance
        '
        'Records 0 and 1 are synthetic worksheet-style Description and Year
        'records.  They intentionally have no values for the first two JV rows.
        'Painting them with normal VGrid cell borders makes them look like
        'editable fields.  Paint the two blanks as one neutral white area and
        'remove their internal vertical separators.  The first real JV record
        'is still painted normally and therefore retains a clear starting edge.
        '--------------------------------------------------------------
        If JVRowTag IsNot Nothing AndAlso
           (String.Equals(JVRowTag.RowKind,
                          "JVName",
                          StringComparison.OrdinalIgnoreCase) OrElse
            String.Equals(JVRowTag.RowKind,
                          "OpeningBalance",
                          StringComparison.OrdinalIgnoreCase)) AndAlso
           (e.RecordIndex = 0 OrElse e.RecordIndex = 1) Then

            e.Cache.FillRectangle(Color.White, e.Bounds)

            'Remove the separator immediately to the right of the row caption
            'and the separator between the two synthetic Description/Year blank
            'records.  Do NOT erase the right edge of record 1: that is the useful
            'boundary before the first real Joint Venture value column.
            If e.RecordIndex = 0 Then

                QueueJointVentureSeparatorSuppression(
                    Grid,
                    e.Row.Name & "|JVBlank|0|L",
                    e.Bounds.Left,
                    e.Bounds.Top + 1,
                    e.Bounds.Bottom - 1,
                    Color.White)

                QueueJointVentureSeparatorSuppression(
                    Grid,
                    e.Row.Name & "|JVBlank|0|R",
                    e.Bounds.Right,
                    e.Bounds.Top + 1,
                    e.Bounds.Bottom - 1,
                    Color.White)

            End If

            'Keep only the horizontal row definition.  This gives the blank
            'area the appearance of merged worksheet whitespace without making
            'the JV Name / Opening Balance rows visually run into one another.
            Using HorizontalPen As New Pen(Color.LightGray, 1)

                e.Cache.DrawLine(
                    HorizontalPen,
                    New Point(e.Bounds.Left, e.Bounds.Bottom - 1),
                    New Point(e.Bounds.Right, e.Bounds.Bottom - 1))

            End Using

            e.Handled = True
            Return

        End If

        Dim Binding As JointVentureCellBinding =
            GetJointVentureBinding(
                Grid,
                e.Row,
                e.RecordIndex)

        'Synthetic Description/Year blanks have no workbook binding. Draw them
        'as ordinary white non-editable cells rather than as data cells.
        If Binding Is Nothing Then

            e.Appearance.BackColor = Color.White
            e.Appearance.ForeColor = AbovoBlue
            Return

        End If

        If IsJointVentureCellLocked(
            Grid,
            e.Row,
            e.RecordIndex,
            Binding) Then

            'Use exactly the same visual language as the generic VGrid rule
            'handler used by Funding.
            e.Appearance.BackColor = Color.Lavender
            e.Appearance.ForeColor = AbovoBlue
            Return

        End If

        Dim SourceCell As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets(Binding.SourceSheet).
                Cells(Binding.SourceAddress)

        'For active cells retain the worksheet appearance where useful, just as
        'the generic VGrid does after its read-only/rule checks.
        If SourceCell IsNot Nothing Then
            e.Appearance.BackColor = SourceCell.Fill.BackgroundColor
            e.Appearance.ForeColor = SourceCell.Font.Color
        End If

        If e.CellValue IsNot Nothing AndAlso IsNumeric(e.CellValue) Then

            If CDbl(e.CellValue) < 0 Then

                e.Appearance.ForeColor = Color.Red

                Dim TextToDraw As String = e.CellText

                If Not String.IsNullOrEmpty(TextToDraw) Then

                    If TextToDraw.StartsWith("-") Then
                        TextToDraw = TextToDraw.Substring(1)
                    End If

                    If Not TextToDraw.StartsWith("(") Then
                        TextToDraw = "(" & TextToDraw & ")"
                    End If

                    e.CellText = TextToDraw

                End If

            End If

        End If

        'Do not set Handled. DevExpress still performs the normal cell paint.

    End Sub

    Private Sub JointVenture_ShowingEditor(ByVal sender As Object, ByVal e As CancelEventArgs)
        Dim Grid As VGridControl = TryCast(sender, VGridControl)
        If Grid Is Nothing OrElse Grid.FocusedRow Is Nothing Then Return

        Dim FocusedEditorRow As EditorRow = TryCast(Grid.FocusedRow, EditorRow)
        If FocusedEditorRow Is Nothing Then
            e.Cancel = True
            Return
        End If

        Dim Binding As JointVentureCellBinding =
            GetJointVentureBinding(
                Grid,
                FocusedEditorRow,
                Grid.FocusedRecord)

        If IsJointVentureCellLocked(
            Grid,
            FocusedEditorRow,
            Grid.FocusedRecord,
            Binding) Then

            e.Cancel = True

        End If
    End Sub

    Private Sub JointVenture_CellValueChanged(ByVal sender As Object,
                                               ByVal e As DevExpress.XtraVerticalGrid.Events.CellValueChangedEventArgs)
        Dim Grid As VGridControl = TryCast(sender, VGridControl)
        If Grid Is Nothing OrElse e.Row Is Nothing Then Return

        Dim ChangedEditorRow As EditorRow = TryCast(e.Row, EditorRow)
        If ChangedEditorRow Is Nothing Then Return

        Dim Key As String = ChangedEditorRow.Properties.FieldName & "|" & e.RecordIndex.ToString
        Dim Map As Dictionary(Of String, JointVentureCellBinding) = Nothing
        If Not JointVentureBindings.TryGetValue(Grid, Map) Then Return

        Dim Binding As JointVentureCellBinding = Nothing
        If Not Map.TryGetValue(Key, Binding) OrElse Binding Is Nothing Then Return

        If IsJointVentureCellLocked(
            Grid,
            ChangedEditorRow,
            e.RecordIndex,
            Binding) Then Return

        Dim SourceCell As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets(Binding.SourceSheet).Cells(Binding.SourceAddress)

        Dim DCE As New DataChangeEvent With {
            .ModelID = ModelID,
            .Description = "Joint Venture assumption updated",
            .WSName = Binding.SourceSheet,
            .CellAddress = Binding.SourceAddress,
            .OriginalValue = SourceCell.DisplayText,
            .ChangedValue = If(e.Value Is Nothing OrElse Convert.IsDBNull(e.Value), Nothing, e.Value),
            .DataFormat = Binding.DataFormat,
            .TimeStamp = Now(),
            .UserName = Environment.UserName
        }

        ChangeMan.ProcessChange(DCE)
        UpdateAllRules()
    End Sub

    Private Function GetVGridColumnTag(ByVal Row As BaseRow,
                                       Optional ByVal CellIndex As Integer = 0) As DataColumnTag

        If Row Is Nothing Then Return Nothing

        Dim DirectTag As DataColumnTag =
            TryCast(Row.Tag, DataColumnTag)

        If DirectTag IsNot Nothing Then Return DirectTag

        Dim MultiTag As VGridMultiEditorRowTag =
            TryCast(Row.Tag, VGridMultiEditorRowTag)

        If MultiTag IsNot Nothing Then
            Return MultiTag.GetColumnTag(CellIndex)
        End If

        Return Nothing

    End Function

    Private Function GetVGridColumnIndex(ByVal Row As BaseRow,
                                         Optional ByVal CellIndex As Integer = 0) As Integer

        If Row Is Nothing Then Return -1

        Dim FieldName As String = Nothing

        Dim ER As EditorRow =
            TryCast(Row, EditorRow)

        If ER IsNot Nothing Then

            FieldName = ER.Properties.FieldName

        Else

            Dim MER As MultiEditorRow =
                TryCast(Row, MultiEditorRow)

            If MER IsNot Nothing AndAlso
               CellIndex >= 0 AndAlso
               CellIndex < MER.PropertiesCollection.Count Then

                FieldName =
                    MER.PropertiesCollection(CellIndex).FieldName

            End If

        End If

        If String.IsNullOrEmpty(FieldName) Then Return -1

        If FieldName.StartsWith("Col_", StringComparison.OrdinalIgnoreCase) Then

            Dim Result As Integer

            If Integer.TryParse(FieldName.Substring(4), Result) Then
                Return Result
            End If

        End If

        Return -1

    End Function

    Private Sub VGrid_CustomDrawCell(
    ByVal sender As Object,
    ByVal e As DevExpress.XtraVerticalGrid.Events.CustomDrawRowValueCellEventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If e.Row Is Nothing Then Return

        Dim ColTag As DataColumnTag =
            GetVGridColumnTag(e.Row, e.CellIndex)

        If ColTag Is Nothing Then Return

        Dim ColIndex As Integer =
            GetVGridColumnIndex(e.Row, e.CellIndex)

        If ColIndex < 0 Then Return
        If e.RecordIndex < 0 Then Return

        Dim UBS As AbovoUnboundSource = TryCast(VG.DataSource, AbovoUnboundSource)

        If UBS Is Nothing Then Return

        Dim DSIndex As Integer = UBS.UBSTag.DSIndex

        If DSIndex < 0 OrElse DSIndex >= DataPres.DataSets.Count Then Return

        If e.RecordIndex >= DataPres.DataSets(DSIndex).DataRows.Count Then Return
        If ColIndex >= DataPres.DataSets(DSIndex).DataColumns.Count Then Return

        Dim CellHandle As CellDataPoint =
        DataPres.DataSets(DSIndex).DataRows(e.RecordIndex).DataCells(ColIndex)

        If CellHandle Is Nothing Then Return


        '----------------------------------------------------------
        ' Dummy column
        '----------------------------------------------------------
        If ColTag.IsDummyColumn Then

            e.Appearance.BackColor = Color.White
            e.Appearance.ForeColor = Color.White

            Return

        End If


        '----------------------------------------------------------
        ' Calculated / read-only cells
        '----------------------------------------------------------
        If ColTag.IsCalculated OrElse ColTag.IsReadOnly Then

            e.Appearance.BackColor = Color.WhiteSmoke
            e.Appearance.ForeColor = AbovoBlue

            Return

        End If


        '----------------------------------------------------------
        ' Rules / locked cells
        '----------------------------------------------------------
        If ColTag.HasRules AndAlso CellHandle.IsLocked Then

            e.Appearance.BackColor = Color.Lavender
            e.Appearance.ForeColor = AbovoBlue

            Return

        End If


        '----------------------------------------------------------
        ' Spreadsheet-derived appearance
        '----------------------------------------------------------
        e.Appearance.BackColor = CellHandle.BGColor
        e.Appearance.ForeColor = CellHandle.FoColor


        '----------------------------------------------------------
        ' Negative values in red
        '
        ' Unlike the XtraGrid version, we can simply alter the
        ' appearance and text and allow VGrid to perform its own
        ' normal painting.
        '----------------------------------------------------------
        If e.CellValue IsNot Nothing AndAlso IsNumeric(e.CellValue) Then

            If CDbl(e.CellValue) < 0 Then

                e.Appearance.ForeColor = Color.Red

                Dim StrToWrite As String = e.CellText

                If Not String.IsNullOrEmpty(StrToWrite) Then

                    If StrToWrite.StartsWith("-") Then

                        StrToWrite = StrToWrite.Substring(1)

                    End If

                    If Not StrToWrite.StartsWith("(") Then

                        StrToWrite = "(" & StrToWrite & ")"

                    End If

                    e.CellText = StrToWrite

                End If

            End If

        End If


        '----------------------------------------------------------
        ' Empty-value pseudo masks
        '----------------------------------------------------------
        If String.IsNullOrEmpty(e.CellText) Then

            Select Case ColTag.ShowDefaultmask

                Case 0
                    e.CellText = "£ ,   "

                Case 1
                    e.CellText = "£ .  "

                Case 3
                    e.CellText = " .  %"

            End Select

        End If

        'IMPORTANT:
        'Don't set e.Handled = True here.
        '
        'We've changed Appearance/CellText and allow DevExpress to
        'paint the cell normally.

    End Sub

    Private Sub VGrid_ValidatingEditor(
    ByVal sender As Object,
    ByVal e As DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs)

        If e.Value Is Nothing Then

            e.Valid = True
            Return

        End If

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If VG.FocusedRow Is Nothing Then Return

        Dim FocusedCellIndex As Integer = 0

        Dim FocusedMultiRow As MultiEditorRow =
            TryCast(VG.FocusedRow, MultiEditorRow)

        If FocusedMultiRow IsNot Nothing AndAlso
           VG.ActiveEditor IsNot Nothing Then

            For MultiCellIndex As Integer =
                0 To FocusedMultiRow.PropertiesCollection.Count - 1

                If Object.ReferenceEquals(
                    FocusedMultiRow.PropertiesCollection(MultiCellIndex).RowEdit,
                    VG.ActiveEditor.Properties) Then

                    FocusedCellIndex = MultiCellIndex
                    Exit For

                End If

            Next

        End If

        Dim ColTag As DataColumnTag =
            GetVGridColumnTag(
                VG.FocusedRow,
                FocusedCellIndex)

        If ColTag Is Nothing Then Return


        Select Case ColTag.DataType

            Case "S"

                e.Valid = True
                Return


            Case Else

                If Not IsNumeric(e.Value) Then

                    e.ErrorText =
                    ColTag.ColumnHeading & " must be numeric."

                    e.Valid = False
                    Return

                End If


                Dim NewValue As Double = CDbl(e.Value)


                If ColTag.MaxVal IsNot Nothing Then

                    If NewValue > CDbl(ColTag.MaxVal) Then

                        e.ErrorText =
                        ColTag.ColumnHeading &
                        " cannot exceed " &
                        ColTag.MaxVal.ToString &
                        "."

                        e.Valid = False
                        Return

                    End If

                End If


                If ColTag.MinVal IsNot Nothing Then

                    If NewValue < CDbl(ColTag.MinVal) Then

                        e.ErrorText =
                        ColTag.ColumnHeading &
                        " must be " &
                        ColTag.MinVal.ToString &
                        " or greater."

                        e.Valid = False
                        Return

                    End If

                End If


                e.Valid = True

        End Select

    End Sub

    Private Sub VGrid_ShowingEditor(
    ByVal sender As Object,
    ByVal e As CancelEventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If VG.FocusedRow Is Nothing Then Return


        'MultiEditorRow constituent items have their own ReadOnly/RowEdit
        'settings.  The ShowingEditor event itself does not expose CellIndex,
        'so do not apply the single-row column lookup to these rows here.
        If TypeOf VG.FocusedRow Is MultiEditorRow Then
            Return
        End If

        Dim ColTag As DataColumnTag =
            GetVGridColumnTag(VG.FocusedRow)

        If ColTag Is Nothing Then

            e.Cancel = True
            Return

        End If


        '----------------------------------------------------------
        ' Calculated/read-only rows
        '----------------------------------------------------------
        If ColTag.IsCalculated OrElse ColTag.IsReadOnly Then

            e.Cancel = True
            Return

        End If


        Dim UBS As AbovoUnboundSource =
        TryCast(VG.DataSource, AbovoUnboundSource)

        If UBS Is Nothing Then Return


        Dim DSIndex As Integer = UBS.UBSTag.DSIndex

        If DSIndex < 0 OrElse DSIndex >= DataPres.DataSets.Count Then
            e.Cancel = True
            Return
        End If

        Dim RowIndex As Integer = VG.FocusedRecord

        Dim ColIndex As Integer =
        GetVGridColumnIndex(VG.FocusedRow)


        If RowIndex < 0 OrElse ColIndex < 0 Then

            e.Cancel = True
            Return

        End If


        Dim ThisDataSet As DataCellRange =
        DataPres.DataSets(DSIndex)


        If RowIndex >= ThisDataSet.DataRows.Count Then

            e.Cancel = True
            Return

        End If


        Dim SourceDataRow = ThisDataSet.DataRows(RowIndex)


        '----------------------------------------------------------
        ' Control/spacer rows
        '
        'These are DATA records in your model, despite rows being
        'displayed vertically.
        '----------------------------------------------------------
        If SourceDataRow.IsControlRow Then

            e.Cancel = True
            Return

        End If


        If SourceDataRow.IsSpacerRow Then

            e.Cancel = True
            Return

        End If


        Dim SourceDataPoint As CellDataPoint =
        SourceDataRow.DataCells(ColIndex)


        If SourceDataPoint Is Nothing Then

            e.Cancel = True
            Return

        End If


        '----------------------------------------------------------
        ' Workbook protection always wins, whether or not this column also
        ' carries presentation rules.
        '----------------------------------------------------------
        If SourceDataPoint.IsLocked Then

            e.Cancel = True
            Return

        End If


        '----------------------------------------------------------
        ' No rules means normal editing after the protection check.
        '----------------------------------------------------------
        If Not ColTag.HasRules Then Return


        '----------------------------------------------------------
        ' Spreadsheet fill rule
        '----------------------------------------------------------
        If Me.ActiveSpreadsheet.Range(
        SourceDataPoint.SourceAddress).Fill.PatternType <>
        PatternType.Solid Then

            e.Cancel = True
            Return

        End If


    End Sub

    Private Sub VGrid_ShownEditor(
    ByVal sender As Object,
    ByVal e As EventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If VG.ActiveEditor Is Nothing Then Return


        If TypeOf VG.ActiveEditor Is DevExpress.XtraEditors.CalcEdit Then

            If DblClickCell Then

                Dim CalcEditor As CalcEdit =
                DirectCast(VG.ActiveEditor, CalcEdit)

                CalcEditor.ShowPopup()

                DblClickCell = False

            End If

        End If

    End Sub

    Private Sub VGrid_CustomCellEditor(
    ByVal sender As Object,
    ByVal e As DevExpress.XtraVerticalGrid.Events.GetCustomRowCellEditEventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If e.Row Is Nothing Then Return


        Dim ColTag As DataColumnTag =
            GetVGridColumnTag(e.Row, e.CellIndex)

        If ColTag Is Nothing Then Return


        Dim UBS As AbovoUnboundSource =
        TryCast(VG.DataSource, AbovoUnboundSource)

        If UBS Is Nothing Then Return


        Dim DSIndex As Integer = UBS.UBSTag.DSIndex

        If DSIndex < 0 OrElse DSIndex >= DataPres.DataSets.Count Then Return

        Dim DataRowIndex As Integer = e.RecordIndex


        If DataRowIndex < 0 Then Return

        If DataRowIndex >=
        DataPres.DataSets(DSIndex).DataRows.Count Then

            Return

        End If


        Dim SourceRow =
        DataPres.DataSets(DSIndex).DataRows(DataRowIndex)


        'Equivalent of:
        '
        'If ViewTag.DataSet.DataRows(e.RowHandle).IsSpacerRow Then
        '    e.RepositoryItem = Nothing
        'End If

        If SourceRow.IsSpacerRow Then

            e.RepositoryItem = Nothing
            Return

        End If


        'Normally do nothing else here.
        '
        'The standard editor has already been assigned with:
        '
        '    EditorRow.Properties.RowEdit
        '
        'during VGrid construction.

    End Sub

    Private Sub VGrid_CellEditorForEditing(
        ByVal sender As Object,
        ByVal e As DevExpress.XtraVerticalGrid.Events.GetCustomRowCellEditEventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing OrElse e.Row Is Nothing Then Return

        Dim ColTag As DataColumnTag =
            GetVGridColumnTag(e.Row, e.CellIndex)

        If ColTag Is Nothing Then Return

        'A repository/combo assigned to the row takes precedence.
        If ColTag.HasComboEdit Then Return

        Dim UBS As AbovoUnboundSource = TryCast(VG.DataSource, AbovoUnboundSource)

        If UBS Is Nothing Then Return

        Dim DSIndex As Integer = UBS.UBSTag.DSIndex

        If DSIndex < 0 OrElse DSIndex >= DataPres.DataSets.Count Then Return

        If e.RecordIndex < 0 OrElse
           e.RecordIndex >= DataPres.DataSets(DSIndex).DataRows.Count Then

            Return

        End If

        Dim SourceRow = DataPres.DataSets(DSIndex).DataRows(e.RecordIndex)

        If SourceRow.IsSpacerRow OrElse SourceRow.IsControlRow Then

            e.RepositoryItem = Nothing
            Return

        End If

        'Use preconfigured repository items; do not mutate repository
        'settings from CustomRecordCellEditForEditing.
        Select Case ColTag.DataType

            Case "S", "P"

                'Use the row's normal editor.

            Case "I", "Y"

                e.RepositoryItem = CustCalcEditInteger

            Case Else

                e.RepositoryItem = CustCalcEditDecimal

        End Select

    End Sub

#End Region

#Region "Draw Handlers"
    Private Sub acCustomDrawAcElement(ByVal sender As Object, ByVal e As CustomDrawElementEventArgs)

        If e.Element.Style = ElementStyle.Item Then

            Dim BBounds As Rectangle = e.ObjectInfo.HeaderBounds
            BBounds.X += (4 * DefaultGridCellPadding)
            e.DrawImage()
            e.DrawContextButtons()
            e.DrawExpandCollapseButton()
            e.DrawHeaderBackground()

            e.Cache.DrawString(e.ObjectInfo.Text, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), BBounds, e.Appearance.GetStringFormat())

            Dim pen As Pen = New Pen(AbovoBlue, 2)
            e.Cache.DrawLine(pen, New Point(e.ObjectInfo.HeaderBounds.Left + DefaultGridCellPadding - 2, e.ObjectInfo.HeaderBounds.Top), New Point(e.ObjectInfo.HeaderBounds.Right - DefaultGridCellPadding + 2, e.ObjectInfo.HeaderBounds.Top))
            e.Cache.DrawLine(pen, New Point(e.ObjectInfo.HeaderBounds.Left + DefaultGridCellPadding - 2, e.ObjectInfo.HeaderBounds.Bottom), New Point(e.ObjectInfo.HeaderBounds.Right - DefaultGridCellPadding + 2, e.ObjectInfo.HeaderBounds.Bottom))

            e.Handled = True

        End If

    End Sub

    Private Sub GVCustomDrawFooterCell(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.FooterCellCustomDrawEventArgs)

        If e.Info.DisplayText Is Nothing Then

            e.DefaultDraw()
            Return

        End If

        Dim pen As Pen = New Pen(AbovoBlue, 1)
        Dim StrToWrite As String = ""

        Dim r As New Rectangle With {
            .X = e.Bounds.X,
            .Height = e.Bounds.Height,
            .Y = e.Bounds.Y,
            .Width = e.Bounds.Width - DefaultGridCellPadding
            }

        e.Cache.DrawLine(pen, New Point(e.Bounds.X, e.Bounds.Top), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Top))
        e.Cache.DrawLine(pen, New Point(e.Bounds.X, e.Bounds.Bottom), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Bottom))

        e.Appearance.ForeColor = AbovoBlue
        StrToWrite = e.Info.DisplayText

        If IsNumeric(StrToWrite) Then
            If CDbl(e.Info.DisplayText) < 0 Then

                If Microsoft.VisualBasic.Left(StrToWrite, 1) = "-" Then StrToWrite = Microsoft.VisualBasic.Right(StrToWrite, Len(StrToWrite) - 1)
                e.Appearance.ForeColor = Color.Red
                StrToWrite = "(" & StrToWrite & ")"

            End If
        End If


        e.Appearance.DrawString(e.Cache, StrToWrite, r)
        e.Handled = True

    End Sub
    Private Sub GridView_CustomDrawCell(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        Dim GV As Object = sender

        Dim ViewTag As GridViewTag = GV.Tag

        Dim PadRect As New Rectangle With {
            .X = e.Bounds.X - DefaultGridCellPadding - 2,
            .Width = e.Bounds.Width + (2 * DefaultGridCellPadding) + 4,
            .Height = e.Bounds.Height + 4 + (2 * DefaultGridCellPadding),
            .Y = e.Bounds.Y - (DefaultGridCellPadding + 2)
            }


        'If ViewTag.DataSet.DataRows(e.RowHandle).IsControlRow Then

        '    If e.Column.Tag.IsControlColumn Then
        '        Dim pen As Pen = New Pen(Color.White, IdealGridRowHeight / 2)
        '        e.Appearance.BackColor = Color.Silver
        '        e.DefaultDraw()
        '        e.Cache.DrawLine(pen, New Point(PadRect.Right, PadRect.Top), New Point(PadRect.Left, PadRect.Top))
        '        e.Cache.DrawLine(pen, New Point(PadRect.Right, PadRect.Bottom), New Point(PadRect.Left, PadRect.Bottom))
        '    Else
        '        e.Appearance.BackColor = Color.White
        '    End If


        '    Return

        'End If



        Dim CellHandle As CellDataPoint = ViewTag.DataSet.DataRows(e.RowHandle).DataCells(e.Column.AbsoluteIndex)

        Dim ColTag As DataColumnTag = e.Column.Tag

        If ColTag.IsDummyColumn Then
            e.Appearance.BackColor = Color.White
            e.DefaultDraw()
            e.Handled = True
            Return
        End If

        If sender.IsCellSelected(e.RowHandle, e.Column) Then

            'If CellHandle.IsLocked Then

            '    e.Appearance.ForeColor = Color.White
            '    Dim br As New SolidBrush(AbovoBlue)
            '    e.Cache.FillRectangle(br, PadRect)
            '    e.Handled = True

            '    Return

            'End If

            e.Appearance.BackColor = Color.Beige
            e.Appearance.ForeColor = Color.Black
            e.DefaultDraw()
            e.Handled = True
            Return

        End If

        If e.Appearance.BackColor = AbovoBlueL3 Then

            e.Appearance.ForeColor = Color.White

            If ColTag.HasRules Then

                If CellHandle.IsLocked Then


                    Dim br As New SolidBrush(AbovoBlueL3)
                    e.Cache.FillRectangle(br, PadRect)
                    e.Handled = True

                    Return

                End If

            End If

            e.DefaultDraw()
            Return

        End If

        If ColTag.HasRules Then

            If CellHandle.IsLocked Then

                Dim pen As Pen = New Pen(Color.Red, 2)
                e.Cache.FillRectangle(Brushes.Lavender, PadRect)
                e.Appearance.ForeColor = Color.White
                e.Appearance.DrawString(e.Cache, e.DisplayText, e.Bounds)

                e.Handled = True

                Return

            End If

        End If

        If ColTag.IsCalculated OrElse ColTag.IsReadOnly Then

            If ColTag.IsReadOnly Then e.Appearance.BackColor = Color.WhiteSmoke
            If ColTag.IsReadOnly Then e.Appearance.ForeColor = AbovoBlue
            e.DefaultDraw()
            e.Handled = True
            Return

        End If

        'Dim Deffont As New Font(e.Appearance.GetFont, FontStyle.Bold)

        If GV.FocusedRowHandle = e.RowHandle Then

            Dim pen As Pen = New Pen(AbovoBlue, 3)
            e.Cache.DrawLine(pen, New Point(PadRect.Right, PadRect.Top), New Point(PadRect.Left, PadRect.Top))
            e.Cache.DrawLine(pen, New Point(PadRect.Right, PadRect.Bottom), New Point(PadRect.Left, PadRect.Bottom))

        End If

        If GV.FocusedColumn.AbsoluteIndex = e.Column.AbsoluteIndex And GV.FocusedRowHandle = e.RowHandle Then

            e.Appearance.BackColor = Color.WhiteSmoke
            e.Appearance.ForeColor = Color.Black

        End If

        If ColTag.IsCalculated Then

            e.Appearance.ForeColor = AbovoBlue

        End If

        Dim strToWrite As String = e.DisplayText

        e.Appearance.BorderColor = CellHandle.BGColor 'e.Appearance.BackColor
        e.Appearance.BackColor = CellHandle.BGColor
        e.Appearance.DrawBackground(e.Cache, PadRect)
        e.Appearance.ForeColor = CellHandle.FoColor

        If e.DisplayText = "" Then
            If ColTag.ShowDefaultmask = 0 Then
                e.Appearance.DrawString(e.Cache, "£ ,   ", PadRect)
            ElseIf ColTag.ShowDefaultmask = 1 Then
                e.Appearance.DrawString(e.Cache, "£ .  ", PadRect)
            ElseIf ColTag.ShowDefaultmask = 3 Then
                e.Appearance.DrawString(e.Cache, " .  %", PadRect)
            End If
            e.Handled = True
            Return
        End If

        If IsNumeric(e.CellValue) Then

            If e.CellValue < 0 Then

                If Microsoft.VisualBasic.Left(strToWrite, 1) = "-" Then strToWrite = Microsoft.VisualBasic.Right(strToWrite, Len(strToWrite) - 1)

                e.Appearance.ForeColor = Color.Red

                If Microsoft.VisualBasic.Left(strToWrite, 1) <> "(" Then strToWrite = "(" & strToWrite & ")"

            End If

        End If

        DrawCellBorder(e, CellHandle.BGColor)

        e.Appearance.DrawString(e.Cache, strToWrite, e.Bounds)

        e.Handled = True

    End Sub


#End Region

#Region "LiveGrid workbook presentation"

    Private Sub ConfigureLiveGridView(ByVal View As GridView,
                                      ByVal LiveDataSet As DataCellRange)

        If View Is Nothing OrElse LiveDataSet Is Nothing Then Return

        Dim ViewTag As GridViewTag = TryCast(View.Tag, GridViewTag)
        If ViewTag Is Nothing Then Return

        Dim Worksheet As DevExpress.Spreadsheet.Worksheet =
            ExcelModels(ModelID).WB.Worksheets(LiveDataSet.SourceWorksheet)

        Dim SourceRanges As New List(Of DevExpress.Spreadsheet.CellRange)

        If LiveDataSet.LiveGridSourceAreaReferences IsNot Nothing AndAlso
           LiveDataSet.LiveGridSourceAreaReferences.Count > 0 Then

            For Each AreaReference As String In LiveDataSet.LiveGridSourceAreaReferences
                SourceRanges.Add(Worksheet.Range(AreaReference))
            Next
        Else
            SourceRanges.Add(Worksheet.Range(LiveDataSet.DataRange))
        End If

        Dim SourceRange As DevExpress.Spreadsheet.CellRange = SourceRanges(0)

        ViewTag.LiveGridSourceRows = New List(Of Integer)
        ViewTag.LiveGridSourceColumns = New List(Of Integer)

        For RowIndex As Integer = SourceRange.TopRowIndex To SourceRange.BottomRowIndex
            If Worksheet.Rows(RowIndex).Visible Then
                ViewTag.LiveGridSourceRows.Add(RowIndex)
            End If
        Next

        For Each Area As DevExpress.Spreadsheet.CellRange In SourceRanges
            For ColumnIndex As Integer = Area.LeftColumnIndex To Area.RightColumnIndex
                If Worksheet.Columns(ColumnIndex).Visible Then
                    ViewTag.LiveGridSourceColumns.Add(ColumnIndex)
                End If
            Next
        Next

        ViewTag.LiveGridSourceRows.RemoveAll(
            Function(RowIndex) IsLiveProjectionRowBlank(
                Worksheet,
                RowIndex,
                ViewTag.LiveGridSourceColumns))

        View.BeginUpdate()

        Try
            With View
                .OptionsBehavior.AllowAddRows = DefaultBoolean.False
                .OptionsBehavior.AllowDeleteRows = DefaultBoolean.False
                .OptionsBehavior.Editable = False
                .OptionsCustomization.AllowColumnMoving = False
                .OptionsCustomization.AllowFilter = False
                .OptionsCustomization.AllowGroup = False
                .OptionsCustomization.AllowSort = False
                .OptionsCustomization.AllowQuickHideColumns = False
                .OptionsMenu.EnableColumnMenu = False
                .OptionsMenu.EnableFooterMenu = False
                .OptionsSelection.EnableAppearanceFocusedRow = False
                .OptionsSelection.MultiSelect = True
                .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
                .OptionsClipboard.CopyColumnHeaders = DefaultBoolean.False
                .OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.None
                .OptionsView.ColumnAutoWidth = False
                .OptionsView.ColumnHeaderAutoHeight = DefaultBoolean.True
                .OptionsView.ShowAutoFilterRow = False
                .OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
                .OptionsView.ShowGroupPanel = False
                .OptionsView.ShowIndicator = False
                .OptionsView.RowAutoHeight = False
                .OptionsView.ShowHorizontalLines = DefaultBoolean.True
                .OptionsView.ShowVerticalLines = DefaultBoolean.True
                .RowHeight = 22
                .ColumnPanelRowHeight = 72
                .UserCellPadding = New System.Windows.Forms.Padding(4, 1, 4, 1)
                .VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Never
            End With

            View.PopulateColumns()

            Dim PresentedColumnCount As Integer =
                Math.Min(
                    Math.Min(View.Columns.Count, LiveDataSet.DataColumns.Length),
                    ViewTag.LiveGridSourceColumns.Count)

            For ColumnOffset As Integer = 0 To PresentedColumnCount - 1
                Dim GridColumn As DevExpress.XtraGrid.Columns.GridColumn =
                    View.Columns(ColumnOffset)

                Dim ColumnTag As DataColumnTag =
                    LiveDataSet.DataColumns(ColumnOffset).ColumnTag

                GridColumn.Tag = ColumnTag
                Dim WorkbookCaption As String =
                    BuildLiveGridColumnCaption(
                        Worksheet,
                        SourceRange,
                        ViewTag.LiveGridSourceColumns(ColumnOffset),
                        String.Empty,
                        LiveDataSet.LiveGridHeaderRows)
                GridColumn.Caption =
                    If(String.IsNullOrWhiteSpace(WorkbookCaption),
                       ColumnTag.ColumnHeading,
                       WorkbookCaption)
                GridColumn.Visible =
                    Not (String.IsNullOrWhiteSpace(WorkbookCaption) AndAlso
                         IsLiveProjectionColumnBlank(
                             Worksheet,
                             ViewTag.LiveGridSourceColumns(ColumnOffset),
                             ViewTag.LiveGridSourceRows))
                GridColumn.OptionsColumn.AllowEdit = False
                GridColumn.OptionsColumn.AllowFocus = True
                GridColumn.OptionsColumn.AllowMerge = DefaultBoolean.False
                GridColumn.OptionsColumn.ReadOnly = True
                GridColumn.AppearanceHeader.Options.UseTextOptions = True
                GridColumn.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap
                GridColumn.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
                GridColumn.AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom

                If ColumnOffset < ViewTag.LiveGridSourceColumns.Count Then
                    Dim SourceColumnIndex As Integer =
                        ViewTag.LiveGridSourceColumns(ColumnOffset)

                    GridColumn.MinWidth = 44
                    GridColumn.Width =
                        Math.Max(
                            56,
                            Math.Min(
                                160,
                                CInt(Math.Round(Worksheet.Columns(SourceColumnIndex).Width / 4.0F))))
                End If
            Next

            AddHandler View.CustomColumnDisplayText,
                AddressOf LiveGridCustomColumnDisplayText
            AddHandler View.RowCellStyle,
                AddressOf LiveGridRowCellStyle

        Finally
            View.EndUpdate()
        End Try

    End Sub

    Private Shared Function IsLiveProjectionRowBlank(
        ByVal Worksheet As DevExpress.Spreadsheet.Worksheet,
        ByVal SourceRowIndex As Integer,
        ByVal SourceColumnIndexes As IEnumerable(Of Integer)) As Boolean

        For Each SourceColumnIndex As Integer In SourceColumnIndexes
            If Not String.IsNullOrWhiteSpace(
                Worksheet.Cells(SourceRowIndex, SourceColumnIndex).DisplayText) Then
                Return False
            End If
        Next

        Return True

    End Function

    Private Shared Function IsLiveProjectionColumnBlank(
        ByVal Worksheet As DevExpress.Spreadsheet.Worksheet,
        ByVal SourceColumnIndex As Integer,
        ByVal SourceRowIndexes As IEnumerable(Of Integer)) As Boolean

        For Each SourceRowIndex As Integer In SourceRowIndexes
            If Not String.IsNullOrWhiteSpace(
                Worksheet.Cells(SourceRowIndex, SourceColumnIndex).DisplayText) Then
                Return False
            End If
        Next

        Return True

    End Function

    Private Shared Function IsNegativeWorkbookCell(
        ByVal SourceCell As DevExpress.Spreadsheet.Cell) As Boolean

        If SourceCell.Value.IsNumeric AndAlso SourceCell.Value.NumericValue < 0 Then
            Return True
        End If

        Dim DisplayValue As String = SourceCell.DisplayText.Trim()
        Return DisplayValue.StartsWith("-", StringComparison.Ordinal) OrElse
               DisplayValue.StartsWith("(", StringComparison.Ordinal)

    End Function

    Private Shared Function BuildLiveGridColumnCaption(
        ByVal Worksheet As DevExpress.Spreadsheet.Worksheet,
        ByVal SourceRange As DevExpress.Spreadsheet.CellRange,
        ByVal SourceColumnIndex As Integer,
        ByVal FallbackCaption As String,
        Optional ByVal HeaderRows As String = Nothing) As String

        Dim CaptionParts As New List(Of String)
        Dim HeaderRowIndexes As New List(Of Integer)

        If Not String.IsNullOrWhiteSpace(HeaderRows) Then
            For Each HeaderRowToken As String In HeaderRows.Split(","c)
                Dim OneBasedRow As Integer
                If Integer.TryParse(HeaderRowToken.Trim(), OneBasedRow) AndAlso OneBasedRow > 0 Then
                    HeaderRowIndexes.Add(OneBasedRow - 1)
                End If
            Next
        End If

        'When XML does not explicitly identify heading rows, use the two nearest
        'non-empty workbook cells above this source column. This supports the
        'different spacing used across Workings sheets while keeping formula-driven
        'headings live.
        If HeaderRowIndexes.Count = 0 Then
            For HeaderRowIndex As Integer = SourceRange.TopRowIndex - 1 To Math.Max(0, SourceRange.TopRowIndex - 8) Step -1
                Dim HeaderText As String =
                    Worksheet.Cells(HeaderRowIndex, SourceColumnIndex).DisplayText.Trim()

                If HeaderText.Length > 0 Then
                    HeaderRowIndexes.Insert(0, HeaderRowIndex)
                    If HeaderRowIndexes.Count = 2 Then Exit For
                End If
            Next
        End If

        For Each HeaderRowIndex As Integer In HeaderRowIndexes

            Dim HeaderText As String =
                Worksheet.Cells(HeaderRowIndex, SourceColumnIndex).DisplayText.Trim()

            If HeaderText.Length > 0 AndAlso Not CaptionParts.Contains(HeaderText) Then
                CaptionParts.Add(HeaderText)
            End If
        Next

        If CaptionParts.Count > 0 Then Return String.Join(vbLf, CaptionParts)

        Return If(FallbackCaption, String.Empty).Replace("vblf", vbLf).Trim()

    End Function

    Private Sub RefreshLiveGridHeaders(ByVal Grid As GridControl)

        If Grid Is Nothing OrElse Grid.IsDisposed Then Return

        Dim View As GridView = TryCast(Grid.MainView, GridView)
        If View Is Nothing Then Return

        Dim ViewTag As GridViewTag = TryCast(View.Tag, GridViewTag)

        If ViewTag Is Nothing OrElse
           Not ViewTag.IsLiveGrid OrElse
           ViewTag.LiveGridSourceColumns Is Nothing OrElse
           ViewTag.DataSet Is Nothing OrElse
           ViewTag.DataSet.DataColumns Is Nothing Then Return

        Dim Worksheet As DevExpress.Spreadsheet.Worksheet =
            ExcelModels(ViewTag.ModelID).WB.Worksheets(ViewTag.LiveGridWorksheet)
        Dim SourceRange As DevExpress.Spreadsheet.CellRange =
            Worksheet.Range(ViewTag.LiveGridRange)
        Dim PresentedColumnCount As Integer =
            Math.Min(
                Math.Min(View.Columns.Count, ViewTag.DataSet.DataColumns.Length),
                ViewTag.LiveGridSourceColumns.Count)

        View.BeginUpdate()

        Try
            For ColumnOffset As Integer = 0 To PresentedColumnCount - 1
                Dim GridColumn As GridColumn = View.Columns(ColumnOffset)
                Dim FallbackCaption As String =
                    ViewTag.DataSet.DataColumns(ColumnOffset).ColumnTag.ColumnHeading
                Dim CurrentCaption As String =
                    BuildLiveGridColumnCaption(
                        Worksheet,
                        SourceRange,
                        ViewTag.LiveGridSourceColumns(ColumnOffset),
                        FallbackCaption,
                        ViewTag.DataSet.LiveGridHeaderRows)

                If GridColumn.Caption <> CurrentCaption Then
                    GridColumn.Caption = CurrentCaption
                End If
            Next
        Finally
            View.EndUpdate()
        End Try

    End Sub

    Private Sub RefreshLiveVGridHeaders(ByVal Grid As VGridControl)

        If Grid Is Nothing OrElse Grid.IsDisposed Then Return

        Dim LayoutTag As VGridLayoutTag = TryCast(Grid.Tag, VGridLayoutTag)
        Dim LiveSource As AbovoUnboundSource = TryCast(Grid.DataSource, AbovoUnboundSource)
        If LayoutTag Is Nothing OrElse Not LayoutTag.IsLiveGrid OrElse
           LiveSource Is Nothing OrElse LiveSource.UBSTag Is Nothing Then Return

        Dim LiveTag As AbovoUnboundSourceTag = LiveSource.UBSTag
        Dim DataSet As DataCellRange = DataPres.DataSets(LiveTag.DSIndex)
        If DataSet Is Nothing OrElse DataSet.DataColumns Is Nothing Then Return

        Dim Worksheet As DevExpress.Spreadsheet.Worksheet =
            ExcelModels(LiveTag.ModelID).WB.Worksheets(LiveTag.LiveGridWorksheet)
        Dim SourceRange As DevExpress.Spreadsheet.CellRange = Worksheet.Range(DataSet.DataRange)

        Grid.BeginUpdate()
        Try
            For Each Row As BaseRow In Grid.Rows
                RefreshLiveVGridRowHeader(Row, Worksheet, SourceRange, DataSet)
            Next
            Grid.InvalidateRecordHeaders()
        Finally
            Grid.EndUpdate()
        End Try

    End Sub

    Private Sub RefreshLiveVGridRowHeader(
        ByVal Row As BaseRow,
        ByVal Worksheet As DevExpress.Spreadsheet.Worksheet,
        ByVal SourceRange As DevExpress.Spreadsheet.CellRange,
        ByVal DataSet As DataCellRange)

        Dim ValueTag As LiveVGridRowTag = TryCast(Row.Tag, LiveVGridRowTag)
        If ValueTag IsNot Nothing Then
            Dim ColumnOffset As Integer = ValueTag.SourceColumnIndex - SourceRange.LeftColumnIndex
            Dim FallbackCaption As String = Row.Properties.Caption
            If ColumnOffset >= 0 AndAlso ColumnOffset < DataSet.DataColumns.Length Then
                FallbackCaption = DataSet.DataColumns(ColumnOffset).ColumnTag.ColumnHeading
            End If
            Row.Properties.Caption = BuildLiveGridColumnCaption(
                Worksheet,
                SourceRange,
                ValueTag.SourceColumnIndex,
                FallbackCaption,
                DataSet.LiveGridHeaderRows)
        End If

        Dim CategoryTag As LiveVGridCategoryTag = TryCast(Row.Tag, LiveVGridCategoryTag)
        If CategoryTag IsNot Nothing Then
            Row.Properties.Caption = Worksheet.Cells(
                CategoryTag.SourceRowIndex,
                CategoryTag.SourceColumnIndex).DisplayText.Trim()
        End If

        For Each ChildRow As BaseRow In Row.ChildRows
            RefreshLiveVGridRowHeader(ChildRow, Worksheet, SourceRange, DataSet)
        Next

    End Sub

    Private Function TryGetLiveGridSourceCell(ByVal View As GridView,
                                              ByVal RowHandle As Integer,
                                              ByVal Column As GridColumn,
                                              ByRef SourceCell As DevExpress.Spreadsheet.Cell) As Boolean

        SourceCell = Nothing

        If View Is Nothing OrElse RowHandle < 0 OrElse Column Is Nothing Then Return False

        Dim ViewTag As GridViewTag = TryCast(View.Tag, GridViewTag)

        If ViewTag Is Nothing OrElse
           Not ViewTag.IsLiveGrid OrElse
           ViewTag.LiveGridSourceRows Is Nothing OrElse
           ViewTag.LiveGridSourceColumns Is Nothing Then Return False

        Dim DataRowIndex As Integer = View.GetDataSourceRowIndex(RowHandle)
        Dim DataColumnIndex As Integer = Column.AbsoluteIndex

        If DataRowIndex < 0 OrElse
           DataRowIndex >= ViewTag.LiveGridSourceRows.Count OrElse
           DataColumnIndex < 0 OrElse
           DataColumnIndex >= ViewTag.LiveGridSourceColumns.Count Then Return False

        SourceCell =
            ExcelModels(ViewTag.ModelID).WB.Worksheets(ViewTag.LiveGridWorksheet).Cells(
                ViewTag.LiveGridSourceRows(DataRowIndex),
                ViewTag.LiveGridSourceColumns(DataColumnIndex))

        Return SourceCell IsNot Nothing

    End Function

    Private Sub LiveGridCustomColumnDisplayText(
        ByVal sender As Object,
        ByVal e As CustomColumnDisplayTextEventArgs)

        Dim View As GridView = TryCast(sender, GridView)
        If View Is Nothing OrElse e.ListSourceRowIndex < 0 Then Return

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        Dim RowHandle As Integer = View.GetRowHandle(e.ListSourceRowIndex)

        If TryGetLiveGridSourceCell(View, RowHandle, e.Column, SourceCell) Then
            e.DisplayText = SourceCell.DisplayText
        End If

    End Sub

    Private Sub LiveGridRowCellStyle(ByVal sender As Object,
                                     ByVal e As RowCellStyleEventArgs)

        Dim View As GridView = TryCast(sender, GridView)
        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing

        If Not TryGetLiveGridSourceCell(View, e.RowHandle, e.Column, SourceCell) Then Return

        Dim Background As Color = SourceCell.FillColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White

        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = AbovoBlue
        If IsNegativeWorkbookCell(SourceCell) Then
            Foreground = Color.Red
        End If

        e.Appearance.BackColor = Background
        e.Appearance.ForeColor = Foreground
        e.Appearance.Options.UseBackColor = True
        e.Appearance.Options.UseForeColor = True
        e.Appearance.Options.UseTextOptions = True
        e.Appearance.TextOptions.WordWrap = WordWrap.Wrap

        Select Case SourceCell.Alignment.Horizontal
            Case SpreadsheetHorizontalAlignment.Center
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Center
            Case SpreadsheetHorizontalAlignment.Right
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Far
            Case SpreadsheetHorizontalAlignment.General
                e.Appearance.TextOptions.HAlignment =
                    If(SourceCell.Value.IsNumeric, HorzAlignment.Far, HorzAlignment.Near)
            Case Else
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Near
        End Select

        Select Case SourceCell.Alignment.Vertical
            Case SpreadsheetVerticalAlignment.Top
                e.Appearance.TextOptions.VAlignment = VertAlignment.Top
            Case SpreadsheetVerticalAlignment.Center
                e.Appearance.TextOptions.VAlignment = VertAlignment.Center
            Case Else
                e.Appearance.TextOptions.VAlignment = VertAlignment.Bottom
        End Select

        Dim SourceFontStyle As FontStyle = FontStyle.Regular
        If SourceCell.Font.Bold Then SourceFontStyle = SourceFontStyle Or FontStyle.Bold
        If SourceCell.Font.Italic Then SourceFontStyle = SourceFontStyle Or FontStyle.Italic
        If SourceCell.Font.UnderlineType <> DevExpress.Spreadsheet.UnderlineType.None Then
            SourceFontStyle = SourceFontStyle Or FontStyle.Underline
        End If

        e.Appearance.Font =
            New Font(e.Appearance.Font.FontFamily, e.Appearance.Font.Size, SourceFontStyle)
        e.Appearance.Options.UseFont = True

        If View.IsCellSelected(e.RowHandle, e.Column) Then
            e.Appearance.BackColor = Color.Beige
            e.Appearance.ForeColor = Color.Black
        End If

    End Sub

    Private Sub LiveVGrid_CustomDrawCell(
        ByVal sender As Object,
        ByVal e As DevExpress.XtraVerticalGrid.Events.CustomDrawRowValueCellEventArgs)

        Dim Grid As VGridControl = TryCast(sender, VGridControl)
        Dim LiveSource As AbovoUnboundSource =
            If(Grid Is Nothing, Nothing, TryCast(Grid.DataSource, AbovoUnboundSource))
        Dim ValueTag As LiveVGridRowTag = TryCast(e.Row.Tag, LiveVGridRowTag)

        If LiveSource Is Nothing OrElse LiveSource.UBSTag Is Nothing OrElse
           ValueTag Is Nothing Then Return

        Dim LiveTag As AbovoUnboundSourceTag = LiveSource.UBSTag
        If LiveTag.LiveGridSourceRows Is Nothing OrElse
           e.RecordIndex < 0 OrElse e.RecordIndex >= LiveTag.LiveGridSourceRows.Count Then Return

        Dim Worksheet As DevExpress.Spreadsheet.Worksheet =
            ExcelModels(LiveTag.ModelID).WB.Worksheets(LiveTag.LiveGridWorksheet)
        Dim SourceCell As DevExpress.Spreadsheet.Cell = Worksheet.Cells(
            LiveTag.LiveGridSourceRows(e.RecordIndex),
            ValueTag.SourceColumnIndex)

        Dim Background As Color = SourceCell.FillColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White

        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = AbovoBlue
        If IsNegativeWorkbookCell(SourceCell) Then
            Foreground = Color.Red
        End If

        e.Appearance.BackColor = Background
        e.Appearance.ForeColor = Foreground
        e.Appearance.Options.UseBackColor = True
        e.Appearance.Options.UseForeColor = True
        e.Appearance.Options.UseTextOptions = True
        e.Appearance.TextOptions.WordWrap = WordWrap.Wrap

        Select Case SourceCell.Alignment.Horizontal
            Case SpreadsheetHorizontalAlignment.Center
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Center
            Case SpreadsheetHorizontalAlignment.Right
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Far
            Case SpreadsheetHorizontalAlignment.General
                e.Appearance.TextOptions.HAlignment =
                    If(SourceCell.Value.IsNumeric, HorzAlignment.Far, HorzAlignment.Near)
            Case Else
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Near
        End Select

        Select Case SourceCell.Alignment.Vertical
            Case SpreadsheetVerticalAlignment.Top
                e.Appearance.TextOptions.VAlignment = VertAlignment.Top
            Case SpreadsheetVerticalAlignment.Center
                e.Appearance.TextOptions.VAlignment = VertAlignment.Center
            Case Else
                e.Appearance.TextOptions.VAlignment = VertAlignment.Bottom
        End Select

        Dim SourceFontStyle As FontStyle = FontStyle.Regular
        If SourceCell.Font.Bold Then SourceFontStyle = SourceFontStyle Or FontStyle.Bold
        If SourceCell.Font.Italic Then SourceFontStyle = SourceFontStyle Or FontStyle.Italic
        If SourceCell.Font.UnderlineType <> DevExpress.Spreadsheet.UnderlineType.None Then
            SourceFontStyle = SourceFontStyle Or FontStyle.Underline
        End If
        e.Appearance.Font = New Font(
            e.Appearance.Font.FontFamily,
            e.Appearance.Font.Size,
            SourceFontStyle)
        e.Appearance.Options.UseFont = True

    End Sub

#End Region

#Region "Data and Date Event Handlers"
    Sub SingleCellDirtyMarker(ByVal sender As Object, ByVal e As Object)

        Try

            sender.MarkDirty()

        Catch ex As Exception

        End Try

    End Sub
    Sub SingleCellControlValidatingEditor(sender As Object, e As DevExpress.XtraEditors.Controls.BaseEditValidatingEventArgs)

        If sender.editValue = Nothing Then

            Return

        End If

        Dim DataTag As SingleCellDataTag = sender.tag

        Select Case DataTag.DataType

            Case "S"

                Return

            Case Else

                Dim Val As Double = sender.editValue

                If Not IsNumeric(Val) Then

                    e.ErrorText = DataTag.Label & " must be numeric."
                    e.Cancel = True
                    Return

                End If

                If DataTag.MaxValSet Then

                    If Val > CDbl(DataTag.MaxVal) Then

                        e.ErrorText = DataTag.Label & " cannot exceed " & DataTag.MaxVal.ToString & "."
                        e.Cancel = True
                        Return

                    End If

                End If

                If DataTag.MinValSet Then

                    If Val < CDbl(DataTag.MinVal) Then

                        e.ErrorText = DataTag.Label & " must be " & DataTag.MinVal.ToString & " or greater."
                        e.Cancel = True
                        Return

                    End If

                End If


        End Select

    End Sub
    Private Sub UnboundDS_ValueNeeded(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValueNeededEventArgs)
        DataCallCount += 1
        Dim UDSSender As AbovoUnboundSource = TryCast(sender, AbovoUnboundSource)

        If InterfaceResourcesReleased OrElse
           UDSSender Is Nothing OrElse
           UDSSender.UBSTag Is Nothing Then

            e.Value = Nothing
            Return
        End If

        If UDSSender.UBSTag.IsLiveGrid Then
            Dim LiveTag As AbovoUnboundSourceTag = UDSSender.UBSTag

            If LiveTag.LiveGridSourceRows Is Nothing OrElse
               LiveTag.LiveGridSourceColumns Is Nothing OrElse
               e.RowIndex < 0 OrElse e.RowIndex >= LiveTag.LiveGridSourceRows.Count OrElse
               e.PropertyIndex < 0 OrElse e.PropertyIndex >= LiveTag.LiveGridSourceColumns.Count Then

                e.Value = Nothing
                Return
            End If

            If ExcelModels Is Nothing OrElse
               LiveTag.ModelID < 0 OrElse LiveTag.ModelID >= ExcelModels.Length OrElse
               ExcelModels(LiveTag.ModelID) Is Nothing OrElse
               ExcelModels(LiveTag.ModelID).IsClosing OrElse
               ExcelModels(LiveTag.ModelID).WB Is Nothing Then

                e.Value = Nothing
                Return
            End If

            e.Value = ExcelModels(LiveTag.ModelID).WB.Worksheets(LiveTag.LiveGridWorksheet).Cells(
                LiveTag.LiveGridSourceRows(e.RowIndex),
                LiveTag.LiveGridSourceColumns(e.PropertyIndex)).DisplayText
            Return
        End If

        e.Value = GetDSData(UDSSender.UBSIndex, UDSSender.UBSTag.DSIndex, e.RowIndex, e.PropertyIndex)

    End Sub

    Private Sub vGridControl_CustomUnboundData(ByVal sender As Object, ByVal e As DevExpress.XtraVerticalGrid.Events.CustomDataEventArgs)

        'Dim ColIndex As Integer = e.Row.VisibleIndex

        'If e.IsGetData Then


        '        Dim DP As CellDataPoint = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(PropertyIndex)

        '        If DP Is Nothing Then Return Nothing

        '        Dim DPC As DevExpress.Spreadsheet.Cell = ExcelModels(ModelID).WB.Worksheets(DP.SourceSheet).Cells(DP.SourceAddress)

        '        If DPC.DisplayText = "" Then

        '            If DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType = "S" Then

        '                Return ""

        '            Else

        '                Return Nothing

        '            End If

        '        End If

        '        Dim ColTag As DataColumnTag = DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag

        '        If ColTag.IsDummyColumn Then Return Nothing

        '        Select Case ColTag.DataType

        '            Case "S"

        '                DP.StringValue = DPC.DisplayText
        '                Return DPC.DisplayText
        '                Exit Function

        '            Case "B"

        '                DP.BoolValue = DPC.Value.NumericValue
        '                Return DPC.Value.NumericValue
        '                Exit Function

        '            Case "N", "P", "C", "M", "R", "SM"
        '                DP.RealValue = DPC.Value.NumericValue
        '                Return DPC.Value.NumericValue
        '                Exit Function

        '            Case "I", "Y"

        '                DP.IntValue = DPC.Value.NumericValue
        '                Return CInt(DPC.Value.NumericValue)
        '                Exit Function

        '            Case Else

        '                Return Nothing

        '        End Select


        '        'If Storage.ContainsKey(e.ListSourceRowIndex) Then
        '        '    e.Value = Storage(e.ListSourceRowIndex)
        '        'Else
        '        '    Storage(e.ListSourceRowIndex) = String.Format("Unbound value {0}", e.ListSourceRowIndex)
        '        '    e.Value = Storage(e.ListSourceRowIndex)
        '        'End If



        '    End If

        '    If e.IsSetData Then
        '    'Storage(e.ListSourceRowIndex) = e.Value.ToString()
        'End If


    End Sub
    Private Function GetDSData(ByVal UBSIndex As Integer, SetDSIndex As Integer, ByVal rowIndex As Integer, ByVal PropertyIndex As Integer) As Object

        'If DataPres.DataSets(SetDSIndex).DataRows(rowIndex).IsControlRow = True Then Return Nothing
        'If DataPres.DataSets(SetDSIndex).DataRows(rowIndex).IsSpacerRow = True Then Return Nothing

        If InterfaceResourcesReleased OrElse
           DataPres Is Nothing OrElse
           DataPres.DataSets Is Nothing OrElse
           SetDSIndex < 0 OrElse SetDSIndex >= DataPres.DataSets.Length OrElse
           DataPres.DataSets(SetDSIndex) Is Nothing OrElse
           DataPres.DataSets(SetDSIndex).DataRows Is Nothing OrElse
           rowIndex < 0 OrElse rowIndex >= DataPres.DataSets(SetDSIndex).DataRows.Length OrElse
           DataPres.DataSets(SetDSIndex).DataRows(rowIndex) Is Nothing OrElse
           DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells Is Nothing OrElse
           PropertyIndex >= DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells.Length OrElse
           DataPres.DataSets(SetDSIndex).DataColumns Is Nothing OrElse
           PropertyIndex < 0 OrElse PropertyIndex >= DataPres.DataSets(SetDSIndex).DataColumns.Length OrElse
           DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag Is Nothing Then

            Return Nothing
        End If

        Dim DP As CellDataPoint = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(PropertyIndex)

        If DP Is Nothing Then Return Nothing

        Dim ColTag As DataColumnTag =
            DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag

        'Some legacy/unbound datasets contain deliberately synthetic cells (for
        'example generated row headings or totals) which have a cached typed
        'value but no workbook source address. Do not ask the worksheet
        'collection for a Nothing key; return the dataset value instead.
        If String.IsNullOrWhiteSpace(DP.SourceSheet) OrElse
           String.IsNullOrWhiteSpace(DP.SourceAddress) Then

            Select Case ColTag.DataType
                Case "S"
                    Return If(DP.StringValue, String.Empty)
                Case "B"
                    Return DP.BoolValue
                Case "N", "P", "C", "M", "R", "SM"
                    Return DP.RealValue
                Case "I", "Y", "D"
                    Return DP.IntValue
                Case Else
                    Return Nothing
            End Select
        End If

        If ExcelModels Is Nothing OrElse
           ModelID < 0 OrElse ModelID >= ExcelModels.Length OrElse
           ExcelModels(ModelID) Is Nothing OrElse
           ExcelModels(ModelID).IsClosing OrElse
           ExcelModels(ModelID).WB Is Nothing OrElse
           Not ExcelModels(ModelID).WB.Worksheets.Contains(DP.SourceSheet) Then

            Return Nothing
        End If

        Dim DPC As DevExpress.Spreadsheet.Cell =
            ExcelModels(ModelID).WB.Worksheets(DP.SourceSheet).Cells(DP.SourceAddress)

        If DPC.DisplayText = "" Then

            If DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType = "S" Then

                Return ""

            Else

                Return Nothing

            End If

        End If

        If ColTag.IsDummyColumn Then Return Nothing

        Select Case ColTag.DataType

            Case "S"

                DP.StringValue = DPC.DisplayText
                Return DPC.DisplayText
                Exit Function

            Case "B"

                DP.BoolValue = DPC.Value.NumericValue
                Return DPC.Value.NumericValue
                Exit Function

            Case "N", "P", "C", "M", "R", "SM"
                DP.RealValue = DPC.Value.NumericValue
                Return DPC.Value.NumericValue
                Exit Function

            Case "I", "Y"

                DP.IntValue = DPC.Value.NumericValue
                Return CInt(DPC.Value.NumericValue)
                Exit Function

            Case Else

                Return Nothing

        End Select

    End Function
    Private Sub UnboundDS_ValuePushed(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValuePushedEventArgs)


        Me.Cursor = Cursors.WaitCursor

        Dim UDSSender As AbovoUnboundSource = sender
        Dim ColSent As Integer = Microsoft.VisualBasic.Right(e.PropertyName, Len(e.PropertyName) - 4)

        PushDSData(UDSSender.UBSTag.DSIndex, e.RowIndex, ColSent, e.Value)

        'On Error Resume Next

        If UDSSender.InBandedMode Then

            UDSSender.ActiveGridBandedView.Columns(ColSent).BestFit()
            UDSSender.ActiveGridBandedView.Columns(ColSent).Width = UDSSender.ActiveGridBandedView.Columns(ColSent).Width * 1.15

        ElseIf UDSSender.InVertMode Then

        Else

            Try
                UDSSender.ActiveGridView.Columns(ColSent).BestFit()
                UDSSender.ActiveGridView.Columns(ColSent).Width = UDSSender.ActiveGridView.Columns(ColSent).Width * 1.15
            Catch ex As Exception

            End Try


        End If

        UpdateRules(UDSSender.UBSTag.DSIndex)

        Me.Cursor = Cursors.Default

    End Sub
    Private Sub PushDSData(ByVal SetDSIndex As Integer, ByVal rowIndex As Integer, ByVal ColSent As Integer, Value As Object, Optional ByVal RefreshAfter As Boolean = True)


        Dim SentRSDT As String = DataPres.DataSets(SetDSIndex).DataColumns(ColSent).ColumnTag.DataType
        Dim SourceDataPoint As CellDataPoint = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent)


        Dim DCE As New DataChangeEvent With {
                    .ModelID = ModelID,
                    .Description = "Cell Updated",
                    .WSName = GetWorkBook(ModelID).Worksheets(SourceDataPoint.SourceSheet).Name,
                    .CellAddress = SourceDataPoint.SourceAddress,
                    .ChangedValue = Value,
                    .DataFormat = SentRSDT,
                    .TimeStamp = Now(),
                    .UserName = Environment.UserName
                }

        Select Case SentRSDT

            Case "S"

                DCE.OriginalValue = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).StringValue

            Case "B"

                DCE.OriginalValue = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).BoolValue.ToString

            Case "N", "P", "C", "D", "M", "R", "SM"

                DCE.OriginalValue = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).RealValue.ToString

            Case "I", "Y"

                DCE.OriginalValue = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).IntValue.ToString

            Case Else
                'DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).StringValue = Value

        End Select

        Dim ChangeResult As AbovoTransaction = ChangeMan.ProcessChange(DCE)

        If ChangeResult.BError Then
            If RefreshAfter Then RefreshData()
            Return
        End If

        DataPres.DataSets(SetDSIndex).IsDirty = True
        Dim WrittenCell As DevExpress.Spreadsheet.Cell =
            GetWorkBook(ModelID).Worksheets(SourceDataPoint.SourceSheet).Cells(SourceDataPoint.SourceAddress)
        SourceDataPoint.IsEmpty = WrittenCell.Value.IsEmpty

        If Not SourceDataPoint.IsEmpty Then
            Select Case SentRSDT
                Case "S"
                    SourceDataPoint.StringValue = WrittenCell.DisplayText
                Case "B"
                    SourceDataPoint.BoolValue = WrittenCell.Value.NumericValue <> 0
                Case "N", "P", "C", "D", "M", "R", "SM"
                    SourceDataPoint.RealValue = WrittenCell.Value.NumericValue
                Case "I", "Y"
                    SourceDataPoint.IntValue = CInt(WrittenCell.Value.NumericValue)
            End Select
        End If

        If RefreshAfter Then RefreshData()

    End Sub
    Sub SingleCell_Value_Push(ByVal sender As Object, ByVal e As Object)

        If sender.editvalue = Nothing Then Exit Sub

        Me.Cursor = Cursors.WaitCursor

        Dim DataTag As SingleCellDataTag = sender.tag
        Dim OldValue As DevExpress.Spreadsheet.CellValue = DataTag.TargetWorksheet.Cells(DataTag.TargetCell).Value
        Dim OldValueString As String = OldValue.ToString

        If sender.editvalue.ToString = OldValueString Then

            Try

                sender.ClearDirtyFlag()
                Me.Cursor = Cursors.Default
                Return

            Catch ex As Exception

            End Try

        End If

        If TryCast(sender, AbovoDESpinEdit) IsNot Nothing Then

            If sender.IsDirty = False Then
                Me.Cursor = Cursors.Default
                Return
            End If

        End If

        If TryCast(sender, AbovoDETextEdit) IsNot Nothing Then

            If sender.IsDirty = False Then
                Me.Cursor = Cursors.Default
                Return
            End If

        End If


        Dim DCM As New DataChangeEvent With {
                        .ModelID = ModelID,
                        .Description = DataTag.Label & " updated from " & OldValueString & " to " & sender.editvalue.ToString,
                        .WSName = DataTag.TargetWorksheet.Name,
                        .CellAddress = DataTag.TargetCell,
                        .ChangedValue = sender.editvalue,
                        .OriginalValue = OldValueString,
                        .DataFormat = DataTag.DataType,
                        .TimeStamp = Now(),
                        .UserName = Environment.UserName
                    }

        If ChangeMan.ProcessChange(DCM).BError = True Then

            Try

                Select Case DataTag.DataType

                    Case "S"

                        sender.editvalue = OldValue.TextValue
                        Me.Cursor = Cursors.Default
                        Return

                    Case "D"

                        sender.editvalue = DateTime.FromOADate(OldValue.NumericValue)
                        Me.Cursor = Cursors.Default
                        Return

                    Case Else

                        sender.editvalue = OldValue.NumericValue
                        Me.Cursor = Cursors.Default
                        Return

                End Select

            Catch ex As Exception

            End Try

        End If

        Try
            sender.ClearDirtyFlag()

        Catch ex As Exception

        End Try

        Me.Cursor = Cursors.Default

    End Sub

#Region "Pasting"


#End Region

#Region "GridEvents"
    Private Sub GridView_ValidatingEditor(sender As Object, e As DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs)

        If e.Value = Nothing Then

            e.Valid = True
            Return

        End If

        Dim view As ColumnView = sender
        Dim column As GridColumn = If(TryCast(e, EditFormValidateEditorEventArgs)?.Column, view.FocusedColumn)
        Dim ColTag As DataColumnTag = column.Tag

        Select Case ColTag.DataType

            Case "S"
                e.Valid = True
                Return

            Case Else

                If Not IsNumeric(e.Value) Then

                    e.ErrorText = ColTag.ColumnHeading & " must be numeric."
                    e.Valid = False
                    Return

                End If

                If Not ColTag.MaxVal Is Nothing Then

                    If e.Value > CDbl(ColTag.MaxVal) Then

                        e.ErrorText = ColTag.ColumnHeading & " cannot exceed " & ColTag.MaxVal.ToString & "."
                        e.Valid = False
                        Return
                    End If

                End If

                If ColTag.MinVal IsNot Nothing Then

                    If e.Value < CDbl(ColTag.MinVal) Then

                        e.ErrorText = ColTag.ColumnHeading & " must  be " & ColTag.MinVal.ToString & " or greater."
                        e.Valid = False
                        Return

                    End If

                End If

                e.Valid = True

        End Select

    End Sub

    Private Sub GridView_Event_DoubleClick(ByVal CGVSender As GridView, ByVal e As MouseEventArgs)

        'Dim hitInfo As GridHitInfo = CGVSender.CalcHitInfo(e.Location)
        'Dim ColTag As DataColumnTag = hitInfo.Column.Tag

        DblClickCell = True





    End Sub
    Private Sub GridView_Event_SingleClick(ByVal CGVSender As GridView, ByVal e As MouseEventArgs)

        Dim hitInfo As GridHitInfo = CGVSender.CalcHitInfo(e.Location)

        If e.Button = MouseButtons.Left Then

            CGVSender.ShowEditor()
            Return

        End If

    End Sub


    Sub GridView_FocusedColumnChange(ByVal CGVSender As GridView, ByVal e As FocusedColumnChangedEventArgs)

        CGVSender.ShowEditor()

    End Sub

    Sub GridView_CalcBandedGridRowHeight(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowHeightEventArgs)

        Dim GV As Object = sender

        Dim ViewTag As GridViewTag = GV.Tag

        If ViewTag.DataSet.DataRows(e.RowHandle).IsControlRow Then

            e.RowHeight = CInt(1.5 * IdealGridRowHeight)

        Else

            e.RowHeight = IdealGridRowHeight

        End If

    End Sub
    Sub GridView_ColumnWidthChanged(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.ColumnEventArgs)

        Dim view As GridView = TryCast(sender, GridView)
        Dim ViewTag As GridViewTag = view.Tag
        ViewTag.InManualReizeMode = True
        ' e.Column.MinWidth = e.Column.Width

        'Dim ColTag As DataColumnTag = e.Column.Tag
        'If ColTag Is Nothing Then Return
        'If ColTag.IsControlColumn Then
        '    For Each Column In view.Columns
        '        Dim ColT As DataColumnTag = Column.Tag
        '        If ColT IsNot Nothing Then
        '            If Not ColT.IsControlColumn Then
        '                Column.Width = e.Column.Width - (2 * DefaultGridCellPadding)
        '                Exit For
        '            End If
        '        End If
        '    Next
        'End If

    End Sub
    Sub GridView_RowCellEditForEditing(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs)

        Dim view As GridView = TryCast(sender, GridView)

        If view Is Nothing OrElse e.Column Is Nothing Then Return

        Dim ViewTag As GridViewTag = TryCast(view.Tag, GridViewTag)
        Dim ColTag As DataColumnTag = TryCast(e.Column.Tag, DataColumnTag)

        If ViewTag Is Nothing OrElse ColTag Is Nothing Then Return

        If ColTag.HasComboEdit Then Return

        If e.RowHandle < 0 OrElse e.RowHandle >= ViewTag.DataSet.DataRows.Count Then Return

        If ViewTag.DataSet.DataRows(e.RowHandle).IsSpacerRow OrElse
           ViewTag.DataSet.DataRows(e.RowHandle).IsControlRow Then

            e.RepositoryItem = Nothing
            Return

        End If

        Select Case ColTag.DataType

            Case "S", "P"

                'Use the column's normal editor.

            Case "I", "Y"

                e.RepositoryItem = CustCalcEditInteger

            Case Else

                e.RepositoryItem = CustCalcEditDecimal

        End Select

    End Sub
    Private DummyCount As Integer = 0
    Sub GridView_CustomSummaryCalculate(ByVal sender As Object, ByVal e As CustomSummaryEventArgs)

        Dim view As GridView = TryCast(sender, GridView)
        If e.IsTotalSummary Then
            Select Case e.SummaryProcess
                    ' Start calculation
                Case CustomSummaryProcess.Start
                    DummyCount = 0
                    ' Consequent calculations
                Case CustomSummaryProcess.Calculate

                    DummyCount += 1

                Case CustomSummaryProcess.Finalize
                    e.TotalValue = DummyCount
            End Select
        End If


    End Sub

    Sub GridView_CustomCellEditor(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs)

        Dim view As GridView = TryCast(sender, GridView)
        Dim ViewTag As GridViewTag = view.Tag
        Dim ColTag As DataColumnTag = e.Column.Tag

        If ViewTag.DataSet.DataRows(e.RowHandle).IsSpacerRow Then
            e.RepositoryItem = Nothing
            Return
        End If

        'Dim view As GridView = DirectCast(sender, GridView)
        'Dim ViewTag As GridViewTag = view.Tag
        'Dim ColTag As DataColumnTag = view.FocusedColumn.Tag
        'If ColTag.DataType = "S" Then Return
        'If ColTag.IsCalculated Or ColTag.IsReadOnly Then Return
        'If DblClickCell Then
        '    If ColTag.DataType = "I" Then
        '        view.ActiveEditor.Properties = CustCalcEdit.Properties
        '    End If
        '    If ColTag.DataType = "N" Then
        '        view.ActiveEditor.Properties = CalcEdit.Properties
        '    End If
        '    DblClickCell = False
        '    view.ShowEditor()
        'End If
    End Sub

    'Private Sub GridView_CustomRowCellEditForEditing(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs)

    '    Dim view As GridView = DirectCast(sender, GridView)
    '    Dim ViewTag As GridViewTag = view.Tag

    '    If DblClickCell Then

    '        Dim ColTag As DataColumnTag = e.Column.Tag
    '        If ColTag.DataType = "S" Then
    '            DblClickCell = False
    '            Return
    '        End If
    '        If ColTag.IsCalculated Or ColTag.IsReadOnly Then
    '            DblClickCell = False
    '            Return
    '        End If
    '        If ColTag.DataType = "I" Then
    '            e.Column.ColumnEdit = CalcEdit

    '        End If



    '        DblClickCell = False

    '        view.ShowEditor()

    '    End If

    'End Sub
    Private Sub GridView_ShownEditor(ByVal sender As Object, ByVal e As EventArgs)

        Dim gv As GridView = sender

        If TypeOf gv.ActiveEditor Is DevExpress.XtraEditors.CalcEdit Then

            If DblClickCell = True Then

                Dim Clcer As CalcEdit = gv.ActiveEditor
                Clcer.ShowPopup()
                DblClickCell = False


            End If

        End If



    End Sub
    Private Sub GridView_ShowingEditor(ByVal sender As Object, ByVal e As CancelEventArgs)




        Dim view As GridView = DirectCast(sender, GridView)

        Dim ViewTag As GridViewTag = view.Tag

        Dim ColTag As DataColumnTag = view.FocusedColumn.Tag

        If ColTag.IsCalculated Or ColTag.IsReadOnly Then

            e.Cancel = True
            Return

        End If




        If ViewTag.DataSet.DataRows(view.FocusedRowHandle).IsControlRow Then

            e.Cancel = True
            Return

        End If

        Dim SourceDataPoint As CellDataPoint = ViewTag.DataSet.DataRows(view.FocusedRowHandle).DataCells(view.FocusedColumn.AbsoluteIndex)

        If ViewTag.DataSet.DataRows(view.FocusedRowHandle).IsSpacerRow Then


            e.Cancel = True
            Return

        End If
        If SourceDataPoint Is Nothing OrElse SourceDataPoint.IsLocked Then
            e.Cancel = True
            Return
        End If

        If Not ColTag.HasRules Then Return

        If Me.ActiveSpreadsheet.Range(SourceDataPoint.SourceAddress).Fill.PatternType <> PatternType.Solid Then
            e.Cancel = True
            Return

        End If

        'Select Case ColTag.DataType

        '        Case "I", "Y"

        '            DblClickCell = False
        '            CustCalcEdit.Precision = 0
        '            'RepositoryItem = CalcEdit
        '            DblClickCell = False
        '            SendingView.ShowEditor()

        '        Case Else

        '            DblClickCell = False
        '            CustCalcEdit.Precision = 5
        '            'e.RepositoryItem = CalcEdit
        '            DblClickCell = False
        '            SendingView.ShowEditor()
        '            Return

        '    End Select
        'End If

    End Sub

#End Region

    Sub ShowCalcEdit(ByVal SendingView As GridView)

        If SendingView Is Nothing OrElse SendingView.FocusedColumn Is Nothing Then Return

        Dim ColTag As DataColumnTag = TryCast(SendingView.FocusedColumn.Tag, DataColumnTag)

        If ColTag Is Nothing Then Return

        'The actual repository item is selected in
        'GridView_RowCellEditForEditing. This routine only asks the view
        'to activate the editor when required.
        If DblClickCell Then

            DblClickCell = False
            SendingView.ShowEditor()

        End If

    End Sub

    Public Sub RunAction(ActToken As ActionToken)

        Select Case ActToken.ActionType

            Case "NRRIbyCOL", "NRRIByCol", "NRRIBYCOL"

                'RepeatsByNR header '+' buttons use NRRIbyCOL.
                '
                'NRRIbyCOL describes the interface presentation, not necessarily the
                'physical source orientation. Most legacy ranges are vertical and
                'transposed into interface columns; Service Charge and some multi-row
                'ranges are horizontal. Expand the named range on its dominant axis.
                Dim args As New XtraInputBoxArgs()

                args.Caption = "Abovo Summit - Add Columns"
                args.Prompt = "Add how many columns?"
                args.DefaultButtonIndex = 0

                Dim editor As New DevExpress.XtraEditors.TextEdit()

                With editor.Properties

                    .UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                    .MaskSettings.Set(
                        "MaskManagerType",
                        GetType(DevExpress.Data.Mask.NumericMaskManager))
                    .MaskSettings.Set("mask", "n0")

                End With

                args.Editor = editor
                args.DefaultResponse = "5"

                Dim result As Object = XtraInputBox.Show(args)

                If result Is Nothing Then
                    editor.Dispose()
                    Return
                End If

                Dim NewColumns As Integer = 0

                If Not Integer.TryParse(
                    result.ToString,
                    NewColumns) Then

                    XtraMessageBox.Show(
                        "Please enter a whole number.",
                        "Abovo Summit",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                    editor.Dispose()
                    Return

                End If

                If NewColumns <= 0 Then
                    editor.Dispose()
                    Return
                End If

                Dim ChangedWorksheet As String =
                    GetStructuralActionWorksheet(
                        ActToken)

                Try

                    Me.Cursor = Cursors.WaitCursor

                    Dim TargetDefinedName As DevExpress.Spreadsheet.DefinedName =
                        ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(
                            ActToken.ActionStrData1)

                    Dim ExpandsByRows As Boolean =
                        TargetDefinedName IsNot Nothing AndAlso
                        TargetDefinedName.Range IsNot Nothing AndAlso
                        TargetDefinedName.Range.RowCount > TargetDefinedName.Range.ColumnCount

                    Dim InsertResult As AbovoTransaction

                    If ExpandsByRows Then
                        InsertResult =
                            WorkbookManager.InsertRows(
                                ModelID,
                                ActToken.ActionStrData1,
                                NewColumns)
                    Else
                        InsertResult =
                            WorkbookManager.InsertColumns(
                                ModelID,
                                ActToken.ActionStrData1,
                                NewColumns)
                    End If

                    If InsertResult IsNot Nothing AndAlso InsertResult.BError Then

                        XtraMessageBox.Show(
                            InsertResult.StringReturn,
                            "Abovo Summit",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)

                        Return

                    End If

                    TransDBManager.CheckTransDBActions(
                        ModelID,
                        ActToken.ActionStrData1)

                    'Refresh every interface section which depends on the
                    'worksheet changed by the repeated named range. This is the
                    'current lazy-tab-aware refresh path.
                    NotifyStructuralWorksheetChange(
                        ChangedWorksheet)

                    UpdateAllRules()

                Finally

                    Me.Cursor = Cursors.Default
                    editor.Dispose()

                End Try

                Return

            Case "NRRI"

                'Semantic structural rules reuse the established NRRI footer
                'infrastructure.  Legacy NRRI named-range expansion continues
                'through the original code below when ActionStrData1 is not a
                'WorkbookStructureRule ID.
                Dim StructuralRuleID As String = Nothing

                If ExcelModels(ModelID).WorkbookStructureRules IsNot Nothing Then

                    StructuralRuleID =
                        ExcelModels(ModelID).WorkbookStructureRules.ResolveRuleID(
                            ActToken.ActionStrData1)

                End If

                If Not String.IsNullOrWhiteSpace(StructuralRuleID) Then

                    Dim StructureArgs As New XtraInputBoxArgs()

                    StructureArgs.Caption = "Abovo Summit - Add / Delete Lines"
                    StructureArgs.Prompt =
                        "Enter the number of lines to add." &
                        Environment.NewLine &
                        "Enter a negative number to delete lines from the end."

                    StructureArgs.DefaultButtonIndex = 0

                    Dim StructureEditor As New DevExpress.XtraEditors.TextEdit()

                    With StructureEditor.Properties
                        .UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                        .MaskSettings.Set(
                            "MaskManagerType",
                            GetType(DevExpress.Data.Mask.NumericMaskManager))
                        .MaskSettings.Set("mask", "n0")
                    End With

                    StructureArgs.Editor = StructureEditor
                    StructureArgs.DefaultResponse = "5"

                    Dim InputResult As Object = XtraInputBox.Show(StructureArgs)

                    If InputResult Is Nothing Then
                        StructureEditor.Dispose()
                        Return
                    End If

                    Dim LineAdjustment As Integer

                    If Not Integer.TryParse(
                        InputResult.ToString,
                        LineAdjustment) Then

                        XtraMessageBox.Show(
                            "Please enter a whole number.",
                            "Abovo Summit",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                        StructureEditor.Dispose()
                        Return
                    End If

                    If LineAdjustment = 0 Then
                        StructureEditor.Dispose()
                        Return
                    End If

                    Dim StructureResult As AbovoTransaction = Nothing

                    Try

                        If LineAdjustment < 0 Then

                            Dim DeleteCount As Integer =
                                Math.Abs(LineAdjustment)

                            'Validate before displaying the confirmation dialog.
                            'This gives the user the real reason immediately if the
                            'requested deletion would breach the rule minimum.
                            Dim DeleteValidation As AbovoTransaction =
                                ExcelModels(ModelID).WorkbookStructureRules.
                                    ValidateDeleteLastRecords(
                                        StructuralRuleID,
                                        DeleteCount)

                            If DeleteValidation.BError Then

                                If Not String.IsNullOrWhiteSpace(
                                    DeleteValidation.StringReturn) Then

                                    XtraMessageBox.Show(
                                        DeleteValidation.StringReturn,
                                        "Abovo Summit",
                                        MessageBoxButtons.OK,
                                        If(DeleteValidation.EventCancelled,
                                           MessageBoxIcon.Information,
                                           MessageBoxIcon.Error))

                                End If

                                Return
                            End If

                            Dim ConfirmResult As DialogResult =
                                XtraMessageBox.Show(
                                    "Delete the last " &
                                    DeleteCount.ToString &
                                    " line(s)?",
                                    "Confirm deletion",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Warning,
                                    MessageBoxDefaultButton.Button2)

                            If ConfirmResult <> DialogResult.Yes Then Return

                        End If

                        'Keep the busy indication active for the entire mutation AND
                        'interface rebuild. RebuildAllSections can recreate child
                        'controls, so UseWaitCursor is used as well as Cursor.
                        Me.UseWaitCursor = True
                        Me.Cursor = Cursors.WaitCursor
                        Cursor.Current = Cursors.WaitCursor

                        If LineAdjustment > 0 Then

                            StructureResult =
                                ExcelModels(ModelID).WorkbookStructureRules.AddRecords(
                                    StructuralRuleID,
                                    LineAdjustment)

                        Else

                            StructureResult =
                                ExcelModels(ModelID).WorkbookStructureRules.DeleteLastRecords(
                                    StructuralRuleID,
                                    Math.Abs(LineAdjustment))

                        End If

                        If StructureResult IsNot Nothing AndAlso
                           StructureResult.BError Then

                            Me.UseWaitCursor = False
                            Me.Cursor = Cursors.Default
                            Cursor.Current = Cursors.Default

                            If Not String.IsNullOrWhiteSpace(
                                StructureResult.StringReturn) Then

                                XtraMessageBox.Show(
                                    StructureResult.StringReturn,
                                    "Abovo Summit",
                                    MessageBoxButtons.OK,
                                    If(StructureResult.EventCancelled,
                                       MessageBoxIcon.Information,
                                       MessageBoxIcon.Error))

                            End If

                            Return
                        End If

                        'Re-assert immediately before the potentially lengthy UI rebuild.
                        Me.UseWaitCursor = True
                        Me.Cursor = Cursors.WaitCursor
                        Cursor.Current = Cursors.WaitCursor

                        RebuildAllSections()

                        Me.UseWaitCursor = True
                        Me.Cursor = Cursors.WaitCursor
                        Cursor.Current = Cursors.WaitCursor

                        ResizeFonts()
                        UpdateAllRules()

                    Finally

                        Me.UseWaitCursor = False
                        Me.Cursor = Cursors.Default
                        Cursor.Current = Cursors.Default
                        StructureEditor.Dispose()

                    End Try

                    Return
                End If

                If String.IsNullOrWhiteSpace(ActToken.ActionStrData1) Then

                    XtraMessageBox.Show(
                        "This grid's Add Lines action does not contain a valid expansion named range.",
                        "Abovo Summit",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

                    Return
                End If

                Dim args As New XtraInputBoxArgs()

                args.Caption = "Abovo Summit - add lines"
                args.Prompt = "Add how many lines?"
                args.DefaultButtonIndex = 0

                Dim editor As New DevExpress.XtraEditors.TextEdit()

                With editor.Properties

                    .UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                    .MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                    .MaskSettings.Set("mask", "n0")

                End With

                args.Editor = editor
                args.DefaultResponse = "5"

                Dim result As Object = XtraInputBox.Show(args)

                If result = Nothing Then Return


                Dim NewRows As Integer = CInt(result)

                If NewRows > 0 Then

                    Me.Cursor = Cursors.WaitCursor

                    Dim ChangedWorksheet As String = GetStructuralActionWorksheet(ActToken)

                    WorkbookManager.InsertRows(ModelID, ActToken.ActionStrData1, NewRows)

                    TransDBManager.CheckTransDBActions(ModelID, ActToken.ActionStrData1)

                    NotifyStructuralWorksheetChange(ChangedWorksheet)
                    UpdateAllRules()

                    Me.Cursor = Cursors.Default

                End If



                editor.Dispose()
                editor = Nothing

            Case "NRCI"

                Dim args As New XtraInputBoxArgs()

                args.Caption = "Abovo Summit - Add Columns"
                args.Prompt = "Add how many columns?"
                args.DefaultButtonIndex = 0

                Dim editor As New DevExpress.XtraEditors.TextEdit()

                With editor.Properties

                    .UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                    .MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                    .MaskSettings.Set("mask", "n0")

                End With

                args.Editor = editor
                args.DefaultResponse = "5"

                Dim result As Object = XtraInputBox.Show(args)

                If result = Nothing Then Return


                Dim NewCols As Integer = CInt(result)

                If NewCols > 0 Then

                    Me.Cursor = Cursors.WaitCursor

                    Dim ChangedWorksheet As String = GetStructuralActionWorksheet(ActToken)

                    WorkbookManager.InsertColumns(ModelID, ActToken.ActionStrData1, NewCols)

                    TransDBManager.CheckTransDBActions(ModelID, ActToken.ActionStrData1)

                    NotifyStructuralWorksheetChange(ChangedWorksheet)
                    UpdateAllRules()

                    Me.Cursor = Cursors.Default

                End If

                editor.Dispose()
                editor = Nothing

        End Select

    End Sub

    Private Function GetStructuralActionWorksheet(ByVal ActToken As ActionToken) As String

        If ActToken Is Nothing OrElse String.IsNullOrWhiteSpace(ActToken.ActionStrData1) Then Return Nothing
        If ExcelModels(ModelID).InterfaceDependencies Is Nothing Then Return Nothing

        Return ExcelModels(ModelID).InterfaceDependencies.GetNamedRangeWorksheetName(ActToken.ActionStrData1)

    End Function

    Private Sub NotifyStructuralWorksheetChange(ByVal WorksheetName As String)

        If String.IsNullOrWhiteSpace(WorksheetName) OrElse ExcelModels(ModelID).InterfaceDependencies Is Nothing Then
            'If the named range cannot be resolved, retain the old broad rebuild as
            'the safe fallback rather than risk leaving a stale interface alive.
            RebuildAllSections()
            Return
        End If

        ExcelModels(ModelID).InterfaceDependencies.WorksheetStructureChanged(WorksheetName)

    End Sub

    Sub UpdateAllRules()

        If DataPres.DataSets.Count = 0 Then Exit Sub

        Dim x As Integer
        For x = 0 To DataPres.DataSets.Count - 1
            If DataPres.DataSets(x).HasRules Then
                DataPres.DataSets(x).UpdateLocks()
            End If
        Next
    End Sub
    Sub UpdateCalcs(DS As Integer)

        If DataPres.DataSets(DS).HasCalcs Then
            DataPres.DataSets(DS).UpdateCalcs()
        End If

    End Sub
    Sub UpdateRules(DS As Integer)

        If DataPres.DataSets(DS).HasRules Then
            DataPres.DataSets(DS).UpdateLocks()
        End If

    End Sub

#End Region

#Region "Interface Extensions"
    Public Class ScrollRediverter : Implements IMessageFilter

        Private ReadOnly OwnerTabControl As XtraTabControl

        Public Sub New(ByVal TabControl As XtraTabControl)

            OwnerTabControl = TabControl

        End Sub

        Public Function PreFilterMessage(
            ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage

            If m.Msg <> DevExpress.Utils.Drawing.Helpers.MSG.WM_MOUSEWHEEL Then
                Return False
            End If

            If OwnerTabControl Is Nothing OrElse
               OwnerTabControl.IsDisposed OrElse
               Not OwnerTabControl.Visible Then

                Return False

            End If

            Dim RecipientPage As XtraTabPage =
                OwnerTabControl.SelectedTabPage

            If RecipientPage Is Nothing OrElse
               RecipientPage.IsDisposed OrElse
               Not RecipientPage.Visible Then

                Return False

            End If

            Dim MouseScreenPoint As Point =
                System.Windows.Forms.Control.MousePosition

            'Only handle a wheel message when the pointer is actually inside the
            'currently selected interface page. Other forms/interfaces must be
            'allowed to process their own wheel messages.
            Dim PageScreenRectangle As Rectangle

            Try

                PageScreenRectangle =
                    RecipientPage.RectangleToScreen(
                        RecipientPage.ClientRectangle)

            Catch

                Return False

            End Try

            If Not PageScreenRectangle.Contains(MouseScreenPoint) Then
                Return False
            End If

            'Normal XtraGrid retains its own vertical data scrolling.
            '
            'VGrid is deliberately different: VGrids are expanded to at least
            'their full visible-content height, so the XtraTabPage owns vertical
            'scrolling for the whole VGrid interface. This avoids nested vertical
            'scroll ranges and ensures synthetic category footers/Add rows actions
            'remain reachable with the same page scrollbar.
            Dim GridScrollOwner As Control =
                FindGridScrollOwnerAtScreenPoint(
                    RecipientPage,
                    MouseScreenPoint)

            If GridScrollOwner IsNot Nothing AndAlso
               TypeOf GridScrollOwner Is GridControl AndAlso
               Not GridScrollOwner.IsDisposed AndAlso
               GridScrollOwner.IsHandleCreated Then

                NativeMethods.SendMessage(
                    GridScrollOwner.Handle,
                    m.Msg,
                    m.WParam,
                    m.LParam)

                Return True

            End If

            'For a VGrid (or any non-grid area), the selected page owns the wheel.
            'This preserves
            'the useful historical behaviour where the wheel scrolls the interface
            'while the pointer is over labels, blank space or other non-scrollable
            'controls.
            If RecipientPage.IsHandleCreated Then

                NativeMethods.SendMessage(
                    RecipientPage.Handle,
                    m.Msg,
                    m.WParam,
                    m.LParam)

                'The manually-forwarded wheel event has been handled. Returning
                'False here was the old double-delivery bug.
                Return True

            End If

            Return False

        End Function

        Private Shared Function FindGridScrollOwnerAtScreenPoint(
            ByVal ParentControl As Control,
            ByVal ScreenPoint As Point) As Control

            If ParentControl Is Nothing OrElse
               ParentControl.IsDisposed OrElse
               Not ParentControl.Visible Then

                Return Nothing

            End If

            'Search descendants in reverse z-order so the visually top-most grid
            'wins if controls overlap.
            For ControlIndex As Integer =
                ParentControl.Controls.Count - 1 To 0 Step -1

                Dim ChildControl As Control =
                    ParentControl.Controls(ControlIndex)

                If ChildControl Is Nothing OrElse
                   ChildControl.IsDisposed OrElse
                   Not ChildControl.Visible Then

                    Continue For

                End If

                Dim ChildScreenRectangle As Rectangle

                Try

                    ChildScreenRectangle =
                        ChildControl.RectangleToScreen(
                            ChildControl.ClientRectangle)

                Catch

                    Continue For

                End Try

                If Not ChildScreenRectangle.Contains(ScreenPoint) Then
                    Continue For

                End If

                'A DevExpress editor hosted inside a grid should still scroll the
                'grid surface rather than the outer XtraTabPage.
                Dim NestedOwner As Control =
                    FindGridScrollOwnerAtScreenPoint(
                        ChildControl,
                        ScreenPoint)

                If NestedOwner IsNot Nothing Then
                    Return NestedOwner
                End If

                If TypeOf ChildControl Is GridControl OrElse
                   TypeOf ChildControl Is VGridControl Then

                    Return ChildControl

                End If

            Next

            If TypeOf ParentControl Is GridControl OrElse
               TypeOf ParentControl Is VGridControl Then

                Return ParentControl

            End If

            Return Nothing

        End Function

    End Class
#End Region
    Private Function GetInColumnEditorTag(ByVal sender As Object) As Object

        If sender Is Nothing Then Return Nothing

        Dim Editor As DevExpress.XtraEditors.BaseEdit = TryCast(sender, DevExpress.XtraEditors.BaseEdit)
        If Editor IsNot Nothing Then

            If Editor.Tag IsNot Nothing Then Return Editor.Tag
            If Editor.Properties IsNot Nothing AndAlso Editor.Properties.Tag IsNot Nothing Then Return Editor.Properties.Tag

        End If

        Dim Repository As DevExpress.XtraEditors.Repository.RepositoryItem = TryCast(sender, DevExpress.XtraEditors.Repository.RepositoryItem)
        If Repository IsNot Nothing Then Return Repository.Tag

        Try
            Return sender.Tag
        Catch
            Return Nothing
        End Try

    End Function

    Private Sub ColumnHeaderEmbededComboChanged(ByVal sender As Object, ByVal e As EventArgs)

        Dim SourceTag As InColumnEditorTagCombo = TryCast(GetInColumnEditorTag(sender), InColumnEditorTagCombo)

        If SourceTag Is Nothing Then Exit Sub

        Dim NewVal As Object = sender.editvalue

        If sender.editvalue = SourceTag.LastEditorValue Then Return



        Dim DCM As New DataChangeEvent(6, True) With {
                        .ModelID = ModelID,
                        .Description = "Column " & SourceTag.EditingNRIndexPosition.ToString & " updated from " & Convert.ToString(SourceTag.LastEditorValue) & " to " & Convert.ToString(NewVal),
                        .TargetNR = SourceTag.EditingNRName,
                        .TargetNRIndex = SourceTag.EditingNRIndexPosition,
                        .ChangedValue = NewVal,
                        .NROrientation = SourceTag.NROrientation,
                        .OriginalValue = SourceTag.LastEditorValue,
                        .DataFormat = SourceTag.EditorFormat,
                        .TimeStamp = Now(),
                        .UserName = Environment.UserName
                    }
        Try

            If ChangeMan.ProcessChangeByNRAddressing(DCM).BError = True Then

                'Select Case SourceTag.EditorFormat

                '    Case "S"

                '        sender.editvalue = OldValue.TextValue
                '        Me.Cursor = Cursors.Default
                '        Return

                '    Case "D"

                '        sender.editvalue = DateTime.FromOADate(OldValue.NumericValue)
                '        Me.Cursor = Cursors.Default
                '        Return

                '    Case Else

                '        sender.editvalue = OldValue.NumericValue
                '        Me.Cursor = Cursors.Default
                '        Return

                'End Select

            Else

                SourceTag.LastEditorValue = NewVal
                SourceTag.LinkedComboBoxEdit.EditValue = NewVal
                If SourceTag.InPlaceColumnHelper IsNot Nothing Then
                    SourceTag.InPlaceColumnHelper.EditValue = NewVal
                End If

                If SourceTag.InPlaceVGridRowHelper IsNot Nothing Then
                    SourceTag.InPlaceVGridRowHelper.EditValue = NewVal
                End If

                UpdateAllRules()

            End If

        Catch ex As Exception

        End Try



    End Sub
    Private Sub ColumnHeaderEmbededDateEChanged(ByVal sender As Object, ByVal e As EventArgs)

        Dim SourceTag As InColumnEditorTagDateEdit = TryCast(GetInColumnEditorTag(sender), InColumnEditorTagDateEdit)

        If SourceTag Is Nothing Then Exit Sub

        Dim NewVal As Object = sender.editvalue

        Dim RevColumn As Integer = -1

        Dim LastEditorText As String = Convert.ToString(SourceTag.LastEditorValue)
        Dim NewEditorText As String = Convert.ToString(NewVal)

        If LastEditorText = "" Then

            If NewEditorText = "" Then
                Return
            Else
                RevColumn = 1
            End If

        Else

            If sender.editvalue = SourceTag.LastEditorValue Then
                Return
            Else
                RevColumn = 2
            End If

        End If

        Dim DCM As New DataChangeEvent(6, True) With {
                        .ModelID = ModelID,
                        .Description = "Column " & SourceTag.EditingNRIndexPosition.ToString & " updated from " & Convert.ToString(SourceTag.LastEditorValue) & " to " & Convert.ToString(NewVal),
                        .TargetNR = SourceTag.EditingNRName,
                        .TargetNRIndex = SourceTag.EditingNRIndexPosition,
                        .ChangedValue = NewVal,
                        .NROrientation = SourceTag.NROrientation,
                        .OriginalValue = SourceTag.LastEditorValue,
                        .DataFormat = SourceTag.EditorFormat,
                        .TimeStamp = Now(),
                        .UserName = Environment.UserName
                    }
        Try

            If ChangeMan.ProcessChangeByNRAddressing(DCM).BError = True Then

                'Select Case SourceTag.EditorFormat

                '    Case "S"

                '        sender.editvalue = OldValue.TextValue
                '        Me.Cursor = Cursors.Default
                '        Return

                '    Case "D"

                '        sender.editvalue = DateTime.FromOADate(OldValue.NumericValue)
                '        Me.Cursor = Cursors.Default
                '        Return

                '    Case Else

                '        sender.editvalue = OldValue.NumericValue
                '        Me.Cursor = Cursors.Default
                '        Return

                'End Select

            Else

                SourceTag.LastEditorValue = NewVal
                SourceTag.LinkedDateBoxEdit.EditValue = NewVal
                If SourceTag.InPlaceColumnHelper IsNot Nothing Then
                    SourceTag.InPlaceColumnHelper.EditValue = NewVal
                End If

                If SourceTag.InPlaceVGridRowHelper IsNot Nothing Then
                    SourceTag.InPlaceVGridRowHelper.EditValue = NewVal
                End If

                UpdateAllRules()

            End If

        Catch ex As Exception

        End Try



    End Sub


    Sub ResizeFonts()

        'Exit Sub


        Dim HaveDoneGrid As Boolean = False

        If ParentGroupForm Is Nothing Then
            Scalefactor = Me.Width / 1700
        Else
            If ParentGroupForm IsNot Nothing Then
                Scalefactor = ParentGroupForm.ClientSize.Width / 1700.0F
            Else
                Scalefactor = Me.ClientSize.Width / 1700.0F
            End If
        End If


        Dim NewFont As Font = GetFont("Small", Scalefactor)

        For Each control In Me.Controls

            control.Font = NewFont

        Next

        If GridViewCount < 0 Then GoTo BandedGridViews

        If Me.UsedGridVIEWS.Length > 0 Then

            For Each GV In Me.UsedGridVIEWS

                If GV Is Nothing Then Continue For
                GV.CheckLoaded()

                If GV.State = GridState.Editing Then
                    GV.CloseEditor()
                End If

                GV.BeginUpdate()

                For Each ap As AppearanceObject In GV.Appearance

                    ap.Font = NewFont

                Next

                Dim ViewTag As GridViewTag = GV.Tag

                If Not ViewTag.HaveProcessedColumns Then Formatter.ProcessGVColumWidths(GV, Me)

                'If Not HaveDoneGrid Then

                GV.BestFitColumns()
                HaveDoneGrid = True
                Dim GVI As GridViewInfo = GV.GetViewInfo()
                IdealGridRowHeight = GVI.MinRowHeight

                'End If

                GV.EndUpdate()

            Next

        End If

BandedGridViews:

        If BandGridViewsCount < 0 Then GoTo TextBoxes

        If Me.UsedBANDedGridVIEWS.Length > 0 Then

            For Each BGV In Me.UsedBANDedGridVIEWS

                If BGV Is Nothing Then Continue For
                BGV.CheckLoaded()

                BGV.BeginUpdate()

                For Each ap As AppearanceObject In BGV.Appearance

                    ap.Font = NewFont

                Next

                Dim ViewTag As GridViewTag = BGV.Tag

                If Not ViewTag.HaveProcessedColumns Then Formatter.ProcessGVColumWidths(BGV, Me)


                'Formatter.ProcessGVColumWidths(BGV, Me)

                'If Not HaveDoneGrid Then

                BGV.BestFitColumns()
                HaveDoneGrid = True
                Dim GVI As GridViewInfo = BGV.GetViewInfo()
                IdealGridRowHeight = GVI.MinRowHeight

                'End If

                BGV.EndUpdate()

            Next

        End If

TextBoxes:

        If TextEditCount < 0 Then GoTo ADETextEdits

        If Me.TextBoxes.Length > 0 Then

            For Each TB In Me.TextBoxes


                TB.Font = NewFont


            Next

        End If

ADETextEdits:

        If ADETextEditsCount < 0 Then GoTo ADICombos

        If Me.ADETextEdits.Length > 0 Then

            For Each ADETB In Me.ADETextEdits

                ADETB.Font = NewFont

            Next

        End If

ADICombos:

        If ADEComboBoxesCount < 0 Then GoTo Combos



        If Me.ADEComboBoxes.Length > 0 Then
            For Each adicb In Me.ADEComboBoxes
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                adicb.Properties.Appearance.Font = NewFont

                ' Next
            Next
        End If




Combos:
        If CombosCount < 0 Then GoTo ADELabels



        If Me.Combos.Length > 0 Then
            For Each cb In Me.Combos
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                cb.Properties.Appearance.Font = NewFont

                ' Next
            Next
        End If

        'Private ADELabels() As AbovoDELabel
        'Private ADELabelsCount As Integer = -1
ADELabels:

        If ADELabelsCount < 0 Then GoTo Labels

        If Me.ADELabels.Length > 0 Then



            For Each adelb In Me.ADELabels
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                adelb.Font = NewFont

                ' Next
            Next

        End If
Labels:

        If LabelsCount < 0 Then GoTo ADEDateBoxes



        If Me.Labels.Length > 0 Then

            For Each lb In Me.Labels
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                lb.Font = NewFont

                ' Next
            Next

        End If

ADEDateBoxes:

        If ADEDateEditsCount < 0 Then GoTo DateBoxes

        If Me.ADEDateEdits.Length > 0 Then


            For Each ADDEE In Me.ADEDateEdits
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                ADDEE.Properties.Appearance.Font = NewFont

                ' Next
            Next

        End If
DateBoxes:

        If DateEditCount < 0 Then GoTo HLinks



        If Me.DateBoxes.Length > 0 Then

            For Each ADbE In Me.DateBoxes
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                ADbE.Properties.Appearance.Font = NewFont

                ' Next
            Next

        End If
        'Private ADEHyper           links() As AbovoDEHyperlinkLabel
        'Private             As Integer = -1
HLinks:

        If ADEHyperlinksCount < 0 Then Exit Sub


        If Me.ADEHyperlinks.Length > 0 Then



            For Each ADhl In Me.ADEHyperlinks
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                ADhl.Font = NewFont

                ' Next
            Next

        End If

TPans:

        'The old test was inverted and exited precisely when TPs existed, so
        'the TablePanels were never resized by ResizeFonts/ResizeControlsCommand.
        If Me.TPs Is Nothing Then Exit Sub

        If Me.TPs.Length > 0 Then

            Dim AvailableWidth As Integer

            If ParentGroupForm IsNot Nothing Then
                AvailableWidth = Math.Max(1, ParentGroupForm.ClientSize.Width - 40)
            Else
                AvailableWidth = Math.Max(1, Me.ClientSize.Width - 40)
            End If

            For Each tp In Me.TPs

                If tp Is Nothing OrElse tp.IsDisposed Then Continue For

                tp.Width = AvailableWidth
                tp.PerformLayout()

            Next

        End If



        'Me.GridControlStockGrid.Font = GetFont("Small", Me.Scalefactor)

        'Me.GridViewStockNumbers.Appearance.OddRow.Font = GetFont("Small", Me.Scalefactor)
        'Me.GridViewStockNumbers.Appearance.ViewCaption.Font = GetFont("Small", Me.Scalefactor)

        'Dim x As Integer

        'For x = 0 To GridViewStockNumbers.Columns.Count - 1

        '    GridViewStockNumbers.Columns(x).AppearanceCell.Font = GetFont("Small", Me.Scalefactor)
        '    GridViewStockNumbers.Columns(x).AppearanceHeader.Font = GetFont("Small", Me.Scalefactor, True)

        'Next

        'RepositoryItemComboBoxOwnedManaged.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemLookUpEditSOCIStockType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemLookUpEditSOCIRentType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemIntegerEdit.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemComboBoxSOCIStockType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)


        'colPropertyStockDescription1.Width = GridControlStockGrid.Width * 0.25
        'colPropertyOwnedManaged1.Width = GridControlStockGrid.Width * 0.125
        'colPropertySOCIStockType1.Width = GridControlStockGrid.Width * 0.125
        'colPropertySOCIRentType1.Width = GridControlStockGrid.Width * 0.125
        'colPropertyCurrentStockNumbers1.Width = GridControlStockGrid.Width * 0.12
        'colPropertyNewLettings1.Width = GridControlStockGrid.Width * 0.12
        'colPropertyTotalOpeningStockCalc1.Width = GridControlStockGrid.Width * 0.12


        'colPropertyPreBPlanStartDateNewBuild1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateDemolitions1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateRTBs1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateOtherDisposals1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyExistingStocksCalc1.Width = GridControlStockGrid.Width * 0.1

        'colPropertyNewLettings1.Width = GridControlStockGrid.Width * 0.1


        'Me.hideContainerRight.Font = GetFont("Small", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.ActiveTab.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDo                 cking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButtonActive.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaptionActive.Font = GetFont("Medium", Me.ScaleFactor)


        ''XtraTabControlMainNavigator.Appearance.FontSizeDelta = MediumFontSize - XtraTabControlMainNavigator.Appearance.Font.Size
        'Me.BarAndDockingControllerMainScreen.AppearancesBar.Dock.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.AppearancePage.HeaderActive.Font = GetFont("Medium", Me.ScaleFactor, False, True)
        'Me.XtraTabControlMainNavigator.AppearancePage.Header.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.Appearance.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.AppearancePage.HeaderHotTracked.Font = GetFont("Medium", Me.ScaleFactor)

        'WindowsUIButtonPanelExitHelp.Font = GetFont("Small", Me.ScaleFactor)
        'Me.WindowsUIButtonPanelBPActions.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelBPActions.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelBPActions.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelBPActions.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelExitHelp.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelExitHelp.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelExitHelp.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelSaveClose.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelSaveClose.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelSaveClose.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)
        'Me.GroupBoxFileActions.Font = GetFont("Small", Me.ScaleFactor)
        'Me.WindowsUIButtonPanelBPActions.ButtonBackgroundImages

    End Sub
    Sub ResizeControls()

        Dim SetWidth As Integer = Me.Width * 0.17
        ScaleUnits = Me.Width * 0.007

        'PictureBoxAbovoLogo.Top = ScaleUnits
        'PictureBoxAbovoLogo.Left = ScaleUnits
        'PictureBoxAbovoLogo.Width = SetWidth
        'PictureBoxAbovoLogo.Height = CInt(PictureBoxAbovoLogo.Width * 0.483)

        'DockPanelSettings.Width = SetWidth
        'WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'GroupBoxProgramDetails.Width = SetWidth
        'XtraTabControlMainNavigator.Top = ScaleUnits
        'XtraTabControlMainNavigator.Left = PictureBoxAbovoLogo.Right + ScaleUnits
        'XtraTabControlMainNavigator.Width = Me.Width - SetWidth - (5 * ScaleUnits) - hideContainerRight.Width
        'XtraTabControlMainNavigator.Height = Me.Height - (6 * ScaleUnits)
        'XtraTabPageMainHABP.Height = XtraTabControlMainNavigator.PageClientBounds.Height
        'WindowsUIButtonPanelOpenCompare.Left = ScaleUnits
        'WindowsUIButtonPanelOpenCompare.Top = 3 * ScaleUnits
        'GroupBoxFileActions.Left = ScaleUnits
        'GroupBoxFileActions.Top = WindowsUIButtonPanelOpenCompare.Bottom + ScaleUnits
        'GroupBoxFileActions.Width = XtraTabControlMainNavigator.Width - (2 * ScaleUnits)
        'GroupBoxFileActions.Height = XtraTabPageMainHABP.Height - WindowsUIButtonPanelExitHelp.Height - (4 * ScaleUnits)
        'WindowsUIButtonPanelExitHelp.Top = XtraTabControlMainNavigator.Bottom - WindowsUIButtonPanelExitHelp.Height
        'WindowsUIButtonPanelExitHelp.Width = SetWidth
        'WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'WebBrowserBPInfo.Top = (2 * ScaleUnits)
        'WebBrowserBPInfo.Width = GroupBoxFileActions.Width - WindowsUIButtonPanelSaveClose.Width - (3 * ScaleUnits)
        'WebBrowserBPInfo.Height = GroupBoxFileActions.Height - WindowsUIButtonPanelBPActions.Height - (4 * ScaleUnits)
        'WindowsUIButtonPanelSaveClose.Left = WebBrowserBPInfo.Right + ScaleUnits
        'WindowsUIButtonPanelSaveClose.Top = WebBrowserBPInfo.Top
        'WindowsUIButtonPanelSaveClose.Height = GroupBoxFileActions.Height
        'WindowsUIButtonPanelOpenCompare.Width = XtraTabControlMainNavigator.PageClientBounds.Width - (2 * ScaleUnits)
        'WindowsUIButtonPanelBPActions.Top = WebBrowserBPInfo.Bottom + ScaleUnits
        'WindowsUIButtonPanelBPActions.Width = WebBrowserBPInfo.Width
        'GroupBoxProgramDetails.Top = PictureBoxAbovoLogo.Bottom + ScaleUnits
        'GroupBoxProgramDetails.Left = ScaleUnits
        'GroupBoxProgramDetails.Height = WindowsUIButtonPanelExitHelp.Top - PictureBoxAbovoLogo.Bottom - (2 * ScaleUnits)
        'SetBrowserText()

    End Sub

    Public Sub ResizeControlsCommand()

        ResizeFonts()

    End Sub
    Sub AnalyseGrids()



    End Sub


    Class AbovoGridRespoitaryCombo

        Public ID As Integer
        Public RepoistaryID As String
        Public Combo As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        Public GridID As Integer = -1
        Public VGridID As Integer = -1
        Public ModelID As Integer

        Public Sub RefreshDataSource()

            Combo.Items.Clear()
            Dim TempCombo As DevExpress.XtraEditors.Repository.RepositoryItemComboBox = RepositaryItems.GetEditor(RepoistaryID, ModelID).RetCombo

            For Each it In TempCombo.Items

                Combo.Items.Add(it)

            Next

            TempCombo.Dispose()
            TempCombo = Nothing

        End Sub

        Class AbovoTabPage

            Public Index As Integer
            Public TabPage As XtraTabPage
            Public GridAndReposss() As GridAndReposs
            Public GridCountrolRepossIndex As Integer = -1
            Public x As Integer = 0

            Public Sub AddGrid(Grid As GridControl)

                GridCountrolRepossIndex += 1
                ReDim Preserve GridAndReposss(GridCountrolRepossIndex)
                GridAndReposss(GridCountrolRepossIndex) = New GridAndReposs With {.Grid = Grid}

            End Sub

            Public Sub AddVGrid(VertGrid As VGridControl)

                GridCountrolRepossIndex += 1
                ReDim Preserve GridAndReposss(GridCountrolRepossIndex)
                GridAndReposss(GridCountrolRepossIndex) = New GridAndReposs With {.VertGrid = VertGrid}

            End Sub

            Public Sub AddRepCombo(AGR As AbovoGridRespoitaryCombo)

                If GridAndReposss Is Nothing OrElse GridCountrolRepossIndex < 0 Then Exit Sub
                GridAndReposss(GridCountrolRepossIndex).AddRepCombo(AGR)

            End Sub

            Public Sub UpdateReposes()

                If GridAndReposss Is Nothing Then Exit Sub

                For Each GR In GridAndReposss

                    If GR Is Nothing OrElse Not GR.HasRepos Then Continue For

                    If GR.Grid IsNot Nothing AndAlso Not GR.Grid.IsDisposed Then

                        GR.Grid.BeginUpdate()

                        Try
                            If GR.Grid.FocusedView IsNot Nothing Then GR.Grid.FocusedView.BeginUpdate()

                            For Each RepCombo In GR.Reposs
                                If RepCombo IsNot Nothing Then RepCombo.RefreshDataSource()
                            Next

                        Finally
                            If GR.Grid.FocusedView IsNot Nothing Then GR.Grid.FocusedView.EndUpdate()
                            GR.Grid.EndUpdate()
                        End Try

                    ElseIf GR.VertGrid IsNot Nothing AndAlso Not GR.VertGrid.IsDisposed Then

                        GR.VertGrid.BeginUpdate()

                        Try
                            For Each RepCombo In GR.Reposs
                                If RepCombo IsNot Nothing Then RepCombo.RefreshDataSource()
                            Next
                        Finally
                            GR.VertGrid.EndUpdate()
                        End Try

                    End If

                Next

            End Sub

            Class GridAndReposs

                Public Grid As GridControl
                Public VertGrid As VGridControl
                Public Reposs() As AbovoGridRespoitaryCombo
                Private RepossIndex As Integer = -1
                Public HasRepos As Boolean = False

                Public Sub AddRepCombo(AGR As AbovoGridRespoitaryCombo)

                    RepossIndex += 1
                    ReDim Preserve Reposs(RepossIndex)
                    Reposs(RepossIndex) = AGR
                    HasRepos = True

                End Sub

            End Class

        End Class

    End Class

    Private Sub RemoveInterfaceScrollRediverter()

        If InterfaceScrollRediverter Is Nothing Then Exit Sub

        Try

            Application.RemoveMessageFilter(InterfaceScrollRediverter)

        Catch
            'The filter may already have been removed during application shutdown.
        Finally

            InterfaceScrollRediverter = Nothing

        End Try

    End Sub

    Private Sub SuspendLazyTabTransitionRedraw()

        If LazyTabTransitionRedrawDisabled Then Exit Sub
        If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Exit Sub

        Try

            'WM_SETREDRAW=False leaves the last fully-painted interface on screen
            'while the target tab is built. Unlike DrawToBitmap, this does not ask
            'the current control tree to lay itself out or paint in order to create
            'a snapshot, so the complexity of the current page is irrelevant.
            NativeMethods.SendMessage(
                Me.Handle,
                &HB,
                IntPtr.Zero,
                IntPtr.Zero)

            LazyTabTransitionRedrawDisabled = True

        Catch

            LazyTabTransitionRedrawDisabled = False

        End Try

    End Sub

    Private Sub ResumeLazyTabTransitionRedraw()

        If Not LazyTabTransitionRedrawDisabled Then Exit Sub

        Try

            If Not Me.IsDisposed AndAlso Me.IsHandleCreated Then

                NativeMethods.SendMessage(
                    Me.Handle,
                    &HB,
                    New IntPtr(1),
                    IntPtr.Zero)

                'The newly selected page has already had its authoritative layout
                'pass. Repaint the entire form once from that completed state.
                Me.Invalidate(True)
                Me.Update()

            End If

        Catch
        Finally

            LazyTabTransitionRedrawDisabled = False

        End Try

    End Sub

    Private Sub EndLazyTabTransitionUpdate()

        If Not LazyTabTransitionUpdateHeld Then Exit Sub

        Try
            XtraTabControlNewGIT.EndUpdate()
        Catch
        Finally
            LazyTabTransitionUpdateHeld = False
        End Try

    End Sub

    Private Sub DataInterfaceTemplate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        'Do not leave native redraw or BeginUpdate suppressed if the form is
        'closed during a first-time lazy-page transition.
        ResumeLazyTabTransitionRedraw()
        EndLazyTabTransitionUpdate()

        'IMessageFilter registrations are Application-wide and outlive the form
        'unless explicitly removed. Do this BEFORE any early-return paths.
        RemoveInterfaceScrollRediverter()

        If ExcelModels Is Nothing OrElse ModelID < 0 OrElse ModelID >= ExcelModels.Length Then Return
        If ExcelModels(ModelID) Is Nothing OrElse ExcelModels(ModelID).InterfaceDependencies Is Nothing Then Return

        ExcelModels(ModelID).InterfaceDependencies.UnregisterInterface(Me)

    End Sub

    Private Sub XtraTabControlNewGIT_SelectedPageChanging(sender As Object, e As TabPageChangingEventArgs) Handles XtraTabControlNewGIT.SelectedPageChanging

        If SuppressLazyTabEvents Then Exit Sub
        If e.Page Is Nothing Then Exit Sub

        Dim SectionIndex As Integer = XtraTabControlNewGIT.TabPages.IndexOf(e.Page)
        If SectionIndex < 0 Then Exit Sub

        Dim State As InterfaceSectionRuntimeState = Nothing
        Dim NeedsBuild As Boolean = True

        If SectionRuntimeStates.TryGetValue(SectionIndex, State) Then
            NeedsBuild = Not State.IsBuilt OrElse State.IsDirty OrElse State.NeedsPresentationRedefinition
        End If

        If NeedsBuild Then

            'Freeze native painting for the COMPLETE form before any target-page
            'build/layout can disturb the currently selected page. WM_SETREDRAW
            'does not require DrawToBitmap, so it cannot itself provoke a layout
            'or paint of a complex current page.
            SuspendLazyTabTransitionRedraw()

            'Keep the COMPLETE visual transition atomic.
            '
            'BuildSection itself uses BeginUpdate/EndUpdate, but previously that
            'inner lock ended while the OLD page was still selected. Its final
            'layout/repaint could therefore be seen briefly before XtraTabControl
            'performed the actual page switch.
            '
            'Hold one outer BeginUpdate across the page-changing event, the lazy
            'build, the actual tab selection, and the final SelectedPageChanged
            'layout pass. EndUpdate is released only when the new page is ready.
            If Not LazyTabTransitionUpdateHeld Then
                XtraTabControlNewGIT.BeginUpdate()
                LazyTabTransitionUpdateHeld = True
            End If

            Try

                'Build the target while it is still off-screen. The expensive
                'content creation still happens before selection, but the
                'authoritative viewport sizing is deferred until the page is
                'actually selected.
                EnsureSectionBuilt(
                    SectionIndex,
                    True)

                SkipNextRepositoryRefreshSection = SectionIndex

            Catch

                EndLazyTabTransitionUpdate()
                ResumeLazyTabTransitionRedraw()
                Throw

            End Try

        End If

    End Sub

    Private Sub XtraTabControlNewGIT_SelectedPageChanged(sender As Object, e As TabPageChangedEventArgs) Handles XtraTabControlNewGIT.SelectedPageChanged

        If SuppressLazyTabEvents Then Exit Sub
        If e.Page Is Nothing Then Exit Sub

        Dim SectionIndex As Integer = XtraTabControlNewGIT.TabPages.IndexOf(e.Page)
        If SectionIndex < 0 Then Exit Sub

        'Programmatic selection or unusual event ordering can bypass the pre-build,
        'so retain this as a correctness fallback.
        EnsureSectionBuilt(SectionIndex)

        'CRITICAL LAYOUT PASS:
        '
        'SelectedPageChanging deliberately builds a lazy page while it is still
        'OFF-SCREEN.  DevExpress does not give an unselected XtraTabPage its final
        'page-client bounds; the Funding debug showed exactly that:
        '
        '    XtraTabControl Client = 5092 x 1917
        '    target XtraTabPage     = 1258 x 1369
        '
        'Once SelectedPageChanged fires the page is genuinely selected and now has
        'its real viewport.  Re-size grids here, not merely during the off-screen
        'pre-build.
        XtraTabControlNewGIT.PerformLayout()
        e.Page.PerformLayout()

        Dim SelectedTP As TablePanel = Nothing

        If TPs IsNot Nothing AndAlso
           SectionIndex >= 0 AndAlso
           SectionIndex < TPs.Length Then

            SelectedTP = TPs(SectionIndex)

        End If

        If SelectedTP IsNot Nothing AndAlso
           Not SelectedTP.IsDisposed Then

            SelectedTP.PerformLayout()
            ApplySectionFontAndGridLayout(SelectedTP)
            SelectedTP.PerformLayout()

        End If

        'The requested page is now selected, fully laid out and correctly sized.
        '
        'First release DevExpress' nested update lock while native form painting
        'is STILL disabled. Then re-enable WM_SETREDRAW and repaint the form once
        'from the completed selected-page state.
        '
        'This is deliberately stronger than the Fix 41 bitmap overlay: complex
        'pages can contain native child windows which repaint independently, and
        'DrawToBitmap can itself trigger layout/painting. A form-level redraw lock
        'makes the amount of work needed to build the target page irrelevant to
        'what remains visible during that work.
        EndLazyTabTransitionUpdate()
        ResumeLazyTabTransitionRedraw()

#If DEBUG Then
#End If

        If AbovoTabPages Is Nothing OrElse SectionIndex >= AbovoTabPages.Length Then Exit Sub

        If SkipNextRepositoryRefreshSection = SectionIndex Then
            'Repository editors were created from current data during the build itself.
            'Refreshing them immediately again is redundant and contributes to flicker.
            SkipNextRepositoryRefreshSection = -1
            Exit Sub
        End If

        Dim TP As AbovoGridRespoitaryCombo.AbovoTabPage = AbovoTabPages(SectionIndex)
        If TP IsNot Nothing Then TP.UpdateReposes()

    End Sub
    Public Sub UpdateTabPage()
        Dim CurrTab As XtraTabPage = XtraTabControlNewGIT.SelectedTabPage

        EnsureSectionBuilt(XtraTabControlNewGIT.SelectedTabPageIndex)

        If AbovoTabPages Is Nothing Then Exit Sub
        If AbovoTabPages.Length = 0 Then Exit Sub
        For Each TP As AbovoGridRespoitaryCombo.AbovoTabPage In AbovoTabPages
            If TP.TabPage Is CurrTab Then
                TP.UpdateReposes()
                Exit Sub
            End If
        Next
    End Sub

    Private Sub DataInterfaceTemplate_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown


        'The FIRST/initially-selected tab does not necessarily raise
        'SelectedPageChanged during constructor-driven population.  The full
        'Funding trace confirms this: there is no "SELECTED PAGE SIZE" line.
        '
        'Shown is therefore the first reliable point at which:
        '  - the form has a window handle,
        '  - XtraTabControlNewGIT has its final client size,
        '  - the initially-selected XtraTabPage has its real bounds.
        '
        'Run the same authoritative grid/VGrid layout pass here for whichever
        'page is selected when the form first becomes visible.
        Dim SelectedIndex As Integer =
            XtraTabControlNewGIT.SelectedTabPageIndex

        If SelectedIndex >= 0 AndAlso
           SelectedIndex < XtraTabControlNewGIT.TabPages.Count Then

            EnsureSectionBuilt(SelectedIndex)

            Dim SelectedPage As XtraTabPage =
                XtraTabControlNewGIT.TabPages(SelectedIndex)

            XtraTabControlNewGIT.PerformLayout()
            SelectedPage.PerformLayout()

            Dim SelectedTP As TablePanel = Nothing

            If TPs IsNot Nothing AndAlso
               SelectedIndex < TPs.Length Then

                SelectedTP = TPs(SelectedIndex)

            End If

            If SelectedTP IsNot Nothing AndAlso
               Not SelectedTP.IsDisposed Then

                SelectedTP.PerformLayout()
                ApplySectionFontAndGridLayout(SelectedTP)
                SelectedTP.PerformLayout()

            End If

#If DEBUG Then

            If SelectedTP IsNot Nothing Then

                For RowIndex As Integer = 0 To SelectedTP.Rows.Count - 1


                Next

                For Each ChildControl As Control In SelectedTP.Controls


                Next

            End If
#End If

        End If

        'Keep the existing label layout refresh.
        If Labels IsNot Nothing AndAlso Labels.Length > 0 Then

            For Each LB In Labels

                If LB IsNot Nothing AndAlso Not LB.IsDisposed Then
                    LB.PerformLayout()
                End If

            Next

        End If

    End Sub

End Class


