Imports Abovo.DataObject
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports Abovo.PresentationManager
Imports DevExpress.Drawing
Imports DevExpress.Utils
Imports DevExpress.XtraSpreadsheet.Layout.Engine

Namespace Abovo
    Public Class AbovoExtendedDEControls
        Public Class AbovoDESpinEdit

            Inherits DevExpress.XtraEditors.SpinEdit

            Public ModelID As Integer
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public PriorVal As Object
            Public LabelText As String
            Public AdvMode As Boolean = False
            Public IsDirty As Boolean = False
            Public IsReadOnly As Boolean = False
            Public Sub Initialise()

                Dim DataTag As SingleCellDataTag = Tag

                TargetWorksheet = DataTag.TargetWorksheet
                TargetCell = DataTag.TargetCell
                Properties.AllowNullInput = True
                Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Properties.Appearance.ForeColor = Color.Black

                If AdvMode Then

                    Properties.UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                    Properties.AdvancedModeOptions.Label = DataTag.Label
                    Properties.AdvancedModeOptions.LabelAppearance.TextOptions.WordWrap = WordWrap.Wrap

                End If

                Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                EditValue = TargetWorksheet.Cells(TargetCell).Value.NumericValue
                Properties.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                Properties.MaskSettings.Set("mask", "F")

                PriorVal = EditValue

                BackColor = TargetWorksheet.Cells(TargetCell).Fill.BackgroundColor
                ForeColor = TargetWorksheet.Cells(TargetCell).Font.Color

            End Sub
            Public Sub MarkDirty()

                IsDirty = True

            End Sub
            Public Sub ClearDirtyFlag()

                IsDirty = False

            End Sub
            Public Sub RefreshData()

                Try

                    EditValue = TargetWorksheet.Cells(TargetCell).Value.NumericValue

                Catch ex As Exception

                    EditValue = ""

                End Try

                PriorVal = EditValue
                IsDirty = False

                Refresh()

            End Sub

            Protected Overrides Function ProcessCmdKey(
                ByRef msg As System.Windows.Forms.Message,
                ByVal keyData As System.Windows.Forms.Keys) As Boolean

                If TryProcessModelHistoryShortcut(Me, ModelID, keyData) Then Return True
                Return MyBase.ProcessCmdKey(msg, keyData)

            End Function

        End Class
        Public Class AbovoDETextEdit

            Inherits DevExpress.XtraEditors.TextEdit

            Public ModelID As Integer
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public PriorVal As Object
            Private DataTag As SingleCellDataTag
            Public LabelText As String
            Public IsBold As Boolean = False
            Public IsDirty As Boolean = False
            Public AdvMode As Boolean = False
            Public IsReadOnly As Boolean = False
            Public Sub Initialise()

                DataTag = Tag

                TargetWorksheet = DataTag.TargetWorksheet
                TargetCell = DataTag.TargetCell
                Properties.AllowNullInput = True
                Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Properties.Appearance.ForeColor = Color.Black


                If AdvMode Then

                    Properties.UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                    Properties.AdvancedModeOptions.Label = DataTag.Label
                    Properties.AdvancedModeOptions.LabelAppearance.TextOptions.WordWrap = WordWrap.Wrap

                End If

                If IsReadOnly Then

                    Properties.ReadOnly = True

                Else

                    Properties.Appearance.BackColor = Color.LightGray

                End If

                Properties.Mask.UseMaskAsDisplayFormat = False

                Select Case DataTag.DataType

                    Case "S"

                        Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.None
                        Properties.AdvancedModeOptions.ShiftedLabelAppearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                        EditValue = TargetWorksheet.Cells(TargetCell).DisplayText
                    Case "FL"

                        If AdvMode Then Properties.AdvancedModeOptions.ShiftedLabelAppearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                        EditValue = TargetWorksheet.Cells(TargetCell).Value.TextValue

                    Case "B"

                        Properties.AdvancedModeOptions.ShiftedLabelAppearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                        EditValue = TargetWorksheet.Cells(TargetCell).Value.BooleanValue

                    Case "D", "P", "C"

                        Properties.MaskSettings.Set("AutoHideDecimalSeparator", True)
                        Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                        Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                        EditValue = TargetWorksheet.Cells(TargetCell).Value.NumericValue
                        Properties.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                        Properties.MaskSettings.Set("mask", "F")

                    Case "M"

                        Properties.MaskSettings.Set("AutoHideDecimalSeparator", True)
                        Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                        Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                        Properties.DisplayFormat.FormatString = "n0"
                        Properties.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                        Properties.MaskSettings.Set("mask", "c5")
                        EditValue = TargetWorksheet.Cells(TargetCell).Value.NumericValue
                        AddHandler Enter, AddressOf AbovoDETextEdit_GotFocus

                    Case "SM"

                        Properties.MaskSettings.Set("AutoHideDecimalSeparator", True)
                        Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                        Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                        Properties.DisplayFormat.FormatString = "n0"
                        Properties.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                        Properties.MaskSettings.Set("mask", "c3")
                        EditValue = TargetWorksheet.Cells(TargetCell).Value.NumericValue

                    Case "I"

                        Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                        Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                        Properties.DisplayFormat.FormatString = "n0"
                        Properties.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                        Properties.MaskSettings.Set("mask", "D3")
                        EditValue = TargetWorksheet.Cells(TargetCell).Value.NumericValue

                    Case "Y"

                        Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                        Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                        Properties.DisplayFormat.FormatString = "D0"
                        Properties.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                        Properties.MaskSettings.Set("mask", "D0")
                        EditValue = TargetWorksheet.Cells(TargetCell).Value.NumericValue

                End Select

                PriorVal = EditValue

                BackColor = TargetWorksheet.Cells(TargetCell).Fill.BackgroundColor
                ForeColor = TargetWorksheet.Cells(TargetCell).Font.Color
                RefreshData()

            End Sub
            Sub AbovoDETextEdit_GotFocus(sender As Object, e As EventArgs)

                Dim ADETE As AbovoDETextEdit = sender


                If ADETE.EditValue IsNot Nothing Then

                    'ADETE.UpdateDisplayText()

                    BeginInvoke(New MethodInvoker(Sub()
                                                      Dim teststring As String = ADETE.Text
                                                      Dim StartPos As Integer = InStr(teststring, "£")
                                                      Dim EndPos As Integer = InStr(teststring, ".")
                                                      If EndPos = 0 Then EndPos = teststring.Length
                                                      ADETE.SelectionStart = StartPos
                                                      ADETE.SelectionLength = EndPos - StartPos - 1
                                                  End Sub))

                End If
            End Sub
            Public Sub MarkDirty()

                IsDirty = True

            End Sub
            Public Sub ClearDirtyFlag()

                IsDirty = False

            End Sub
            Public Sub RefreshData()

                Try

                    Select Case DataTag.DataType

                        Case "S"

                            EditValue = TargetWorksheet.Cells(TargetCell).DisplayText

                        Case "FL"

                            EditValue = TargetWorksheet.Cells(TargetCell).Value.TextValue

                        Case "B"

                            EditValue = TargetWorksheet.Cells(TargetCell).Value.BooleanValue

                        Case Else

                            EditValue = TargetWorksheet.Cells(TargetCell).Value.NumericValue

                    End Select

                Catch ex As Exception

                    EditValue = ""

                End Try

                PriorVal = EditValue
                IsDirty = False

                Refresh()

            End Sub

            Protected Overrides Function ProcessCmdKey(
                ByRef msg As System.Windows.Forms.Message,
                ByVal keyData As System.Windows.Forms.Keys) As Boolean

                If TryProcessModelHistoryShortcut(Me, ModelID, keyData) Then Return True
                Return MyBase.ProcessCmdKey(msg, keyData)

            End Function

        End Class

        Public Class AbovoDEHyperlinkLabel

            Inherits DevExpress.XtraEditors.HyperlinkLabelControl
            Public ModelID As Integer
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public PriorVal As Object

            Public LoadMode As String = "CR"
            Public IsBold As Boolean = False
            Public IsTitle As Integer = 0
            Public IsStatic As Boolean = False
            Public AutoSizeHorizontal As Boolean = False

            Public Sub Initialise()

                Dim DataTag As ElementInterfaceLinkTag = Tag

                Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near

                AllowHtmlString = True
                LinkBehavior = LinkBehavior.HoverUnderline
                Appearance.LinkColor = AbovoBlue
                Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                Appearance.Options.UseTextOptions = True


                If AutoSizeHorizontal Then
                    AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal
                Else
                    AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Vertical
                End If

                If IsBold Then
                    Appearance.FontSizeDelta = FontStyle.Bold
                End If

                If IsTitle > 0 Then
                    Appearance.Font = New Font(Appearance.Font, FontStyle.Bold)
                    Appearance.FontSizeDelta = 1 + IsTitle
                End If

                RefreshData()

            End Sub
            Public Sub RefreshData()

                If LoadMode = "CR" Then

                    Try

                        Text = TargetWorksheet.Cells(TargetCell).DisplayText

                    Catch ex As Exception

                        Text = "Error"

                    End Try

                End If

                Refresh()

            End Sub

        End Class

        Public Class AbovoDEComboBox

            Inherits DevExpress.XtraEditors.ComboBoxEdit

            Public ModelID As Integer
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public PriorVal As Object
            Public LoadMode As String = "CR"
            Public IsBold As Boolean = False
            Public IsTitle As Integer = 0
            Public IsUnderline As Boolean = False
            Public DoBorder As Boolean = False
            Public RepID As String
            Public IsDirty As Boolean = True
            Public IsStatic As Boolean = False
            Public TxtHlignment As DevExpress.Utils.HorzAlignment = DevExpress.Utils.HorzAlignment.Far
            Public AutoSizeHorizontal As Boolean = False
            Private CurrList As List(Of String)
            Private LitmitToList As Boolean = True

            Public Property SetLimitToList As Boolean
                Get
                    Return LitmitToList
                End Get
                Set(value As Boolean)
                    LitmitToList = value
                    If value Then
                        'Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
                    Else
                        'Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
                    End If
                End Set

            End Property
            Public Sub CommonItems()

                BackColor = TargetWorksheet.Cells(TargetCell).Fill.BackgroundColor
                ForeColor = TargetWorksheet.Cells(TargetCell).Font.Color
                'ForeColor = Color.White

                ProcesDefValue()

                Dim DataTag As SingleCellDataTag = Tag

                Select Case DataTag.DataType

                    Case "S"

                        Properties.AdvancedModeOptions.ShiftedLabelAppearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near

                    Case Else

                        Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

                End Select

            End Sub
            Public Sub RefreshData()

                Try

                    Text = TargetWorksheet.Cells(TargetCell).DisplayText
                    BackColor = TargetWorksheet.Cells(TargetCell).Fill.BackgroundColor
                    'ForeColor = TargetWorksheet.Cells(TargetCell).colo

                Catch ex As Exception

                    MsgBox("Error refreshing label data for cell " & TargetCell & " - " & ex.Message)

                End Try
                IsDirty = False

                Refresh()

            End Sub
            Public Sub MarkDirty()

                IsDirty = True

            End Sub
            Public Sub ClearDirtyFlag()

                IsDirty = False

            End Sub
            Public Sub InitialiseFromNRP(NRName As String)

                Dim DataTag As SingleCellDataTag = Tag

                TargetWorksheet = DataTag.TargetWorksheet
                TargetCell = DataTag.TargetCell

                ClearList()

                Dim ListItems As List(Of String) = RepositaryItems.GetListFromNR(NRName, ModelID)
                Properties.Items.AddRange(ListItems)
                CurrList = ListItems
                CommonItems()

            End Sub
            Public Sub InitialiseStandard(RepID As String)

                ClearList()

                Dim DataTag As SingleCellDataTag = Tag

                TargetWorksheet = DataTag.TargetWorksheet
                TargetCell = DataTag.TargetCell

                Dim ListItems As List(Of String) = RepositaryItems.GetList(RepID, ModelID)
                Properties.Items.AddRange(ListItems)
                CurrList = ListItems

                CommonItems()

            End Sub
            Sub ProcesDefValue()

                Try

                    EditValue = TargetWorksheet.Cells(TargetCell).DisplayText
                    BackColor = TargetWorksheet.Cells(TargetCell).Fill.BackgroundColor

                Catch ex As Exception

                    MsgBox("Error setting default value for combo box - " & ex.Message)

                End Try

            End Sub
            Public Property SetTargetCell As String
                Get
                    Return TargetCell
                End Get
                Set(value As String)
                    TargetCell = value
                End Set
            End Property
            Public Property SetTargetWorksheet As DevExpress.Spreadsheet.Worksheet
                Get
                    Return TargetWorksheet
                End Get
                Set(value As DevExpress.Spreadsheet.Worksheet)
                    TargetWorksheet = value
                End Set
            End Property

            Public Property SetModelID As Integer
                Get
                    Return ModelID
                End Get
                Set(value As Integer)
                    ModelID = value
                End Set
            End Property
            Public Sub ClearList()

                Properties.Items.Clear()

            End Sub
            Protected Sub ProcessChange(ByVal sender As Object, ByVal e As System.EventArgs)

                Dim Result = PostModelCellValue(ModelID, TargetWorksheet.Name, TargetCell,
                                                If(EditValue Is Nothing OrElse Convert.IsDBNull(EditValue), Nothing, EditValue),
                                                "S", "Selection updated")
                If Result.BError Then RefreshData()

            End Sub

            Protected Overrides Function ProcessCmdKey(
                ByRef msg As System.Windows.Forms.Message,
                ByVal keyData As System.Windows.Forms.Keys) As Boolean

                If TryProcessModelHistoryShortcut(Me, ModelID, keyData) Then Return True
                Return MyBase.ProcessCmdKey(msg, keyData)

            End Function

        End Class
        Public Class AbovoDEHeaderDateBox

            Inherits DevExpress.XtraEditors.DateEdit

            Public ModelID As Integer
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public PriorVal As Object
            Public LoadMode As String = "CR"
            Public IsBold As Boolean = False
            Public IsTitle As Integer = 0
            Public IsUnderline As Boolean = False
            Public AddBlankFirstItem As Boolean = False
            Public DoBorder As Boolean = False
            Public RepID As String
            Public IsDirty As Boolean = True
            Public InPlaceColumnHelper As ColumnInplaceEditorHelper
            Public IsStatic As Boolean = False
            Public TxtHlignment As DevExpress.Utils.HorzAlignment = DevExpress.Utils.HorzAlignment.Far
            Public AutoSizeHorizontal As Boolean = False
            Private CurrList As List(Of String)
            Private LitmitToList As Boolean = True



            Public Sub CommonItems()

                BackColor = AbovoComboBGC
                ForeColor = Color.White

            End Sub
            Public Sub RefreshData()

                Try

                    EditValue = TargetWorksheet.Cells(TargetCell).DisplayText

                Catch ex As Exception

                    MsgBox("Error refreshing label data for cell " & TargetCell & " - " & ex.Message)

                End Try
                IsDirty = False

                Refresh()

            End Sub
            Public Sub MarkDirty()

                IsDirty = True

            End Sub
            Public Sub ClearDirtyFlag()

                IsDirty = False

            End Sub
            Public Sub InitialiseFromNRP(NRName As String)

                Dim DataTag As SingleCellDataTag = Tag

                TargetWorksheet = DataTag.TargetWorksheet
                TargetCell = DataTag.TargetCell


                CommonItems()

            End Sub
            Public Sub InitialiseStandard(RepID As String)



                CommonItems()

            End Sub
            Sub ProcesDefValue()

                'Try

                '    EditValue = TargetWorksheet.Cells(TargetCell).DisplayText

                'Catch ex As Exception

                '    MsgBox("Error setting default value for combo box - " & ex.Message)

                'End Try

            End Sub
            Public Property SetTargetCell As String
                Get
                    Return TargetCell
                End Get
                Set(value As String)
                    TargetCell = value
                End Set
            End Property
            Public Property SetTargetWorksheet As DevExpress.Spreadsheet.Worksheet
                Get
                    Return TargetWorksheet
                End Get
                Set(value As DevExpress.Spreadsheet.Worksheet)
                    TargetWorksheet = value
                End Set
            End Property

            Public Property SetModelID As Integer
                Get
                    Return ModelID
                End Get
                Set(value As Integer)
                    ModelID = value
                End Set
            End Property

            Protected Sub ProcessChange(ByVal sender As Object, ByVal e As System.EventArgs)

                Dim Result = PostModelCellValue(ModelID, TargetWorksheet.Name, TargetCell,
                                                If(EditValue Is Nothing OrElse Convert.IsDBNull(EditValue), Nothing, EditValue),
                                                "D", "Date header updated")
                If Result.BError Then RefreshData()

            End Sub

        End Class
        Public NotInheritable Class OrdinalYearComboItem

            Public Sub New(ByVal storedValue As String, ByVal displayText As String)
                Me.StoredValue = storedValue
                Me.DisplayText = displayText
            End Sub

            Public ReadOnly Property StoredValue As String
            Public ReadOnly Property DisplayText As String

            Public Overrides Function ToString() As String
                Return DisplayText
            End Function

            Public Overrides Function Equals(ByVal obj As Object) As Boolean
                Dim other As OrdinalYearComboItem = TryCast(obj, OrdinalYearComboItem)
                If other IsNot Nothing Then
                    Return String.Equals(StoredValue, other.StoredValue, StringComparison.OrdinalIgnoreCase)
                End If

                Return String.Equals(
                    StoredValue,
                    Convert.ToString(obj).Trim(),
                    StringComparison.OrdinalIgnoreCase)
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return StringComparer.OrdinalIgnoreCase.GetHashCode(StoredValue)
            End Function

        End Class

        Public Class AbovoDEHeaderComboBox

            Inherits DevExpress.XtraEditors.ComboBoxEdit

            Public ModelID As Integer
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public PriorVal As Object
            Public LoadMode As String = "CR"
            Public IsBold As Boolean = False
            Public IsTitle As Integer = 0
            Public IsUnderline As Boolean = False
            Public AddBlankFirstItem As Boolean = False
            Public DoBorder As Boolean = False
            Public RepID As String
            Public IsDirty As Boolean = True
            Public InPlaceColumnHelper As ColumnInplaceEditorHelper
            Public IsStatic As Boolean = False
            Public TxtHlignment As DevExpress.Utils.HorzAlignment = DevExpress.Utils.HorzAlignment.Far
            Public AutoSizeHorizontal As Boolean = False
            Private CurrList As List(Of String)
            Private LitmitToList As Boolean = True
            Public RenderedListItems As List(Of String)
            Private OrdinalYearItems As Dictionary(Of String, OrdinalYearComboItem)
            Private OrdinalYearDisplayHandlerAttached As Boolean = False

            Public Property SetLimitToList As Boolean
                Get
                    Return LitmitToList
                End Get
                Set(value As Boolean)
                    LitmitToList = value
                    If value Then
                        'Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
                    Else
                        'Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
                    End If
                End Set

            End Property
            Public Sub CommonItems()

                BackColor = AbovoComboBGC
                ForeColor = Color.White

            End Sub
            Public Sub RefreshData()

                Try

                    Text = TargetWorksheet.Cells(TargetCell).DisplayText

                Catch ex As Exception

                    MsgBox("Error refreshing label data for cell " & TargetCell & " - " & ex.Message)

                End Try
                IsDirty = False

                Refresh()

            End Sub
            Public Sub MarkDirty()

                IsDirty = True

            End Sub
            Public Sub ClearDirtyFlag()

                IsDirty = False

            End Sub
            Public Sub InitialiseFromNRP(NRName As String)

                Dim DataTag As SingleCellDataTag = Tag

                TargetWorksheet = DataTag.TargetWorksheet
                TargetCell = DataTag.TargetCell

                ClearList()

                Dim ListItems As List(Of String) = RepositaryItems.GetListFromNR(NRName, ModelID)
                RenderedListItems = ListItems
                If AddBlankFirstItem Then ListItems.Insert(0, "<Blank>")
                Properties.Items.AddRange(ListItems)
                CurrList = ListItems
                CommonItems()

            End Sub
            Public Sub InitialiseStandard(RepID As String)

                ClearList()
                Me.RepID = RepID

                'Dim DataTag As SingleCellDataTag = Tag

                'TargetWorksheet = DataTag.TargetWorksheet
                'TargetCell = DataTag.TargetCell

                Dim ListItems As List(Of String) = RepositaryItems.GetList(RepID, ModelID)

                If IsOrdinalYearRepository(RepID) Then
                    InitialiseOrdinalYearItems(ListItems)
                Else
                    If AddBlankFirstItem Then ListItems.Insert(0, "<Blank>")
                    RenderedListItems = ListItems
                    Properties.Items.AddRange(ListItems)
                    CurrList = ListItems
                End If

                CommonItems()

            End Sub

            Private Shared Function IsOrdinalYearRepository(ByVal repositoryID As String) As Boolean
                Return String.Equals(repositoryID, "Rep_OrdinalYears", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(repositoryID, "Rep_OrdinalYearsLess1", StringComparison.OrdinalIgnoreCase)
            End Function

            Private Sub InitialiseOrdinalYearItems(ByVal storedValues As IEnumerable(Of String))
                OrdinalYearItems =
                    New Dictionary(Of String, OrdinalYearComboItem)(StringComparer.OrdinalIgnoreCase)

                Dim periodByOrdinal As Dictionary(Of String, String) = LoadOrdinalYearPeriods()
                Dim propertyItems As New List(Of Object)
                Dim displayItems As New List(Of String)

                If AddBlankFirstItem Then
                    propertyItems.Add("<Blank>")
                    displayItems.Add("<Blank>")
                End If

                For Each storedValue As String In storedValues
                    Dim ordinal As String = If(storedValue, String.Empty).Trim()
                    If ordinal.Length = 0 Then Continue For

                    Dim period As String = Nothing
                    Dim displayText As String = ordinal
                    If periodByOrdinal.TryGetValue(ordinal, period) AndAlso
                       Not String.IsNullOrWhiteSpace(period) Then
                        displayText = ordinal & " - " & ShortenYearDescription(period)
                    End If

                    Dim item As New OrdinalYearComboItem(ordinal, displayText)
                    OrdinalYearItems(ordinal) = item
                    propertyItems.Add(item)
                    displayItems.Add(displayText)
                Next

                RenderedListItems = displayItems
                CurrList = displayItems
                Properties.Items.AddRange(propertyItems.ToArray())
                AddHandler Properties.CustomDisplayText, AddressOf OrdinalYear_CustomDisplayText
                OrdinalYearDisplayHandlerAttached = True
            End Sub

            Private Function LoadOrdinalYearPeriods() As Dictionary(Of String, String)
                Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                If FileManager.ExcelModels Is Nothing OrElse
                   ModelID < 0 OrElse ModelID >= FileManager.ExcelModels.Length OrElse
                   FileManager.ExcelModels(ModelID) Is Nothing OrElse
                   FileManager.ExcelModels(ModelID).WB Is Nothing Then
                    Return result
                End If

                Dim yearTable As DevExpress.Spreadsheet.Worksheet = Nothing
                For Each worksheet As DevExpress.Spreadsheet.Worksheet In FileManager.ExcelModels(ModelID).WB.Worksheets
                    If String.Equals(
                        worksheet.Name,
                        "Hidden - Year Table",
                        StringComparison.OrdinalIgnoreCase) Then
                        yearTable = worksheet
                        Exit For
                    End If
                Next

                If yearTable Is Nothing Then Return result

                Dim usedRange As DevExpress.Spreadsheet.CellRange = yearTable.GetUsedRange()
                For rowIndex As Integer = usedRange.TopRowIndex To usedRange.BottomRowIndex
                    Dim ordinal As String = yearTable.Cells(rowIndex, 6).DisplayText.Trim()
                    Dim period As String = yearTable.Cells(rowIndex, 8).DisplayText.Trim()
                    If ordinal.Length > 0 AndAlso period.Length > 0 Then
                        result(ordinal) = period
                    End If
                Next

                Return result
            End Function

            Private Shared Function ShortenYearDescription(ByVal fullDescription As String) As String
                Dim parts As String() = If(fullDescription, String.Empty).Trim().Split("/"c)
                If parts.Length <> 2 Then Return If(fullDescription, String.Empty).Trim()

                Dim firstYear As String = parts(0).Trim()
                Dim secondYear As String = parts(1).Trim()
                If firstYear.Length > 2 Then firstYear = firstYear.Substring(firstYear.Length - 2)
                If secondYear.Length > 2 Then secondYear = secondYear.Substring(secondYear.Length - 2)

                Return firstYear & "/" & secondYear
            End Function

            Private Sub OrdinalYear_CustomDisplayText(
                ByVal sender As Object,
                ByVal e As DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs)

                If OrdinalYearItems Is Nothing OrElse e.Value Is Nothing Then Return

                Dim item As OrdinalYearComboItem = TryCast(e.Value, OrdinalYearComboItem)
                If item IsNot Nothing Then
                    e.DisplayText = item.DisplayText
                    Return
                End If

                Dim ordinal As String = Convert.ToString(e.Value).Trim()
                If OrdinalYearItems.TryGetValue(ordinal, item) Then
                    e.DisplayText = item.DisplayText
                End If
            End Sub
            Sub ProcesDefValue()

                'Try

                '    EditValue = TargetWorksheet.Cells(TargetCell).DisplayText

                'Catch ex As Exception

                '    MsgBox("Error setting default value for combo box - " & ex.Message)

                'End Try

            End Sub
            Public Property SetTargetCell As String
                Get
                    Return TargetCell
                End Get
                Set(value As String)
                    TargetCell = value
                End Set
            End Property
            Public Property SetTargetWorksheet As DevExpress.Spreadsheet.Worksheet
                Get
                    Return TargetWorksheet
                End Get
                Set(value As DevExpress.Spreadsheet.Worksheet)
                    TargetWorksheet = value
                End Set
            End Property

            Public Property SetModelID As Integer
                Get
                    Return ModelID
                End Get
                Set(value As Integer)
                    ModelID = value
                End Set
            End Property
            Public Sub ClearList()

                If OrdinalYearDisplayHandlerAttached Then
                    RemoveHandler Properties.CustomDisplayText, AddressOf OrdinalYear_CustomDisplayText
                    OrdinalYearDisplayHandlerAttached = False
                End If

                OrdinalYearItems = Nothing
                Properties.Items.Clear()

            End Sub
            Protected Sub ProcessChange(ByVal sender As Object, ByVal e As System.EventArgs)

                Dim Result = PostModelCellValue(ModelID, TargetWorksheet.Name, TargetCell,
                                                If(EditValue Is Nothing OrElse Convert.IsDBNull(EditValue), Nothing, EditValue),
                                                "S", "Header selection updated")
                If Result.BError Then RefreshData()

            End Sub

        End Class
        Public Class AbovoDELabel

            Inherits DevExpress.XtraEditors.LabelControl
            Public ModelID As Integer
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public PriorVal As Object
            Public LoadMode As String = "CR"
            Public IsBold As Boolean = False
            Public IsTitle As Integer = 0
            Public IsUnderline As Boolean = False
            Public DoBorder As Boolean = False
            Public IsStatic As Boolean = False
            Public TxtHlignment As DevExpress.Utils.HorzAlignment = DevExpress.Utils.HorzAlignment.Far
            Public AutoSizeHorizontal As Boolean = False

            Public Sub Initialise()

                Dim DataTag As SingleCellDataTag = Tag

                TargetWorksheet = DataTag.TargetWorksheet
                TargetCell = DataTag.TargetCell

                AllowHtmlString = True
                Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                Appearance.Options.UseTextOptions = True

                'BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple

                If DoBorder Then

                    Dim BorderPadding As New Padding
                    BorderPadding.Bottom = 2
                    BorderPadding.Top = 2
                    BorderPadding.Left = 0
                    BorderPadding.Right = 0
                    Padding = BorderPadding

                End If

                If IsUnderline Then

                    Appearance.Font = New Font(Appearance.Font, FontStyle.Underline)

                End If

                If AutoSizeHorizontal Then
                    AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal
                Else
                    AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Vertical
                End If

                If IsBold Then

                    Appearance.FontStyleDelta = FontStyle.Bold

                End If

                If IsTitle Then

                    Appearance.Font = New Font(Appearance.Font, FontStyle.Bold)
                    Appearance.FontSizeDelta = IsTitle

                End If

                RefreshData()

                Appearance.TextOptions.HAlignment = TxtHlignment

                Dim teststring As String = Text.Trim
                teststring = Text.Replace(",", "")
                teststring = Text.Replace(".", "")


                If IsNumeric(teststring) Then
                    Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                    'Else
                    '    Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                End If

                TxtHlignment = Appearance.TextOptions.HAlignment

                RefreshData()

            End Sub
            Public Sub RefreshData()

                If LoadMode = "CR" Then

                    Try

                        Text = TargetWorksheet.Cells(TargetCell).DisplayText

                    Catch ex As Exception

                        MsgBox("Error refreshing label data for cell " & TargetCell & " - " & ex.Message)

                    End Try

                End If

                Refresh()

            End Sub

        End Class

        Public Class AbovoDEDateEdit

            Inherits DevExpress.XtraEditors.DateEdit
            Public ModelID As Integer
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public PriorVal As Object
            Public IsDirty As Boolean = True
            Public LoadMode As String = "CR"
            Public IsStatic As Boolean = False
            Public Sub Initialise()

                Dim DataTag As SingleCellDataTag = Tag

                TargetWorksheet = DataTag.TargetWorksheet
                TargetCell = DataTag.TargetCell

                Try

                    EditValue = DateTime.FromOADate(TargetWorksheet.Range(TargetCell).Value.NumericValue)

                Catch ex As Exception

                    EditValue = ""

                End Try

                PriorVal = EditValue

            End Sub
            Public Sub MarkDirty()

                IsDirty = True

            End Sub
            Public Sub ClearDirtyFlag()

                IsDirty = False

            End Sub
            Public Sub RefreshData()

                Try

                    EditValue = DateTime.FromOADate(TargetWorksheet.Range(TargetCell).Value.NumericValue)

                Catch ex As Exception

                    EditValue = ""

                End Try

                PriorVal = EditValue
                IsDirty = False

                Refresh()

            End Sub

            Protected Overrides Function ProcessCmdKey(
                ByRef msg As System.Windows.Forms.Message,
                ByVal keyData As System.Windows.Forms.Keys) As Boolean

                If TryProcessModelHistoryShortcut(Me, ModelID, keyData) Then Return True
                Return MyBase.ProcessCmdKey(msg, keyData)

            End Function

        End Class

    End Class

End Namespace
