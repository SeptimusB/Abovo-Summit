Imports Abovo.FileManager
Imports Abovo.FontManager

Imports System.ComponentModel.DataAnnotations
Imports System.ServiceModel.Channels
Imports DevExpress.CodeParser
Imports DevExpress.DataAccess.DataFederation
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraRichEdit.Layout
Namespace Abovo
    Public Class RepositaryItems

        Public Repositaries As AbovoRespositaryItem()

        Public Shared Function GetEditorFromList(List As List(Of String)) As AbovoRespositaryItem

            Dim EditorReturn As New AbovoRespositaryItem
            Dim DefaultCombo As New RepositoryItemComboBox

            CType(DefaultCombo, System.ComponentModel.ISupportInitialize).BeginInit()

            With DefaultCombo
                .Appearance.Font = DefaultFont
                .Appearance.Options.UseFont = True
                .AppearanceDropDown.Font = DefaultFont
                .AppearanceDropDown.Options.UseFont = True
                .AutoHeight = False
                .TextEditStyle = TextEditStyles.DisableTextEditor
                .BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
                .Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
                .LookAndFeel.UseDefaultLookAndFeel = False
            End With

            EditorReturn.RepType = "CMB"
            EditorReturn.RetCombo = DefaultCombo
            EditorReturn.ListItems = List
            EditorReturn.RetCombo.Items.AddRange(EditorReturn.ListItems)

            Return EditorReturn

        End Function
        Public Shared Function GetEditor(ID As String, setModelID As Integer, Optional ByVal InsertBlankFirstItem As Boolean = False) As AbovoRespositaryItem

            Dim ModelID As Integer

            ModelID = setModelID
            Dim EditorReturn As New AbovoRespositaryItem
            Dim DefaultCombo As New RepositoryItemComboBox

            CType(DefaultCombo, System.ComponentModel.ISupportInitialize).BeginInit()

            With DefaultCombo
                .Appearance.Font = DefaultFont
                .Appearance.Options.UseFont = True
                .AppearanceDropDown.Font = DefaultFont
                .AppearanceDropDown.Options.UseFont = True
                .AutoHeight = False
                .TextEditStyle = TextEditStyles.DisableTextEditor
                .BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
                .Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
                .LookAndFeel.UseDefaultLookAndFeel = True
            End With

            EditorReturn.RepType = "CMB"
            EditorReturn.RetCombo = DefaultCombo
            Dim ReturnList As New List(Of String)

            Select Case ID

                Case "Rep_OwnMan"

                    ReturnList = New List(Of String)({"Owned", "Managed"})


                Case "Rep_YEByP"

                    ReturnList = New List(Of String)({"Year End", "By Period"})

                Case "Rep_CMCats"

                    ReturnList = GetListFromNR("CMCats", ModelID)

                Case "Rep_ValueBasis"

                    ReturnList = New List(Of String)({"EUV-SH", "Cost", "OMV", "MVT"})

                Case "Rep_IncExclude"

                    ReturnList = New List(Of String)({"Include", "Exclude"})

                Case "Rep_FFRRentType"

                    ReturnList = GetListFromNR("FFRRentType", ModelID)

                Case "Rep_OrdinalYears"

                    Dim OrdYearList As New List(Of String)
                    ' Populate the list with ordinal years from 1 to 40
                    For i As Integer = 1 To 40

                        OrdYearList.Add(i.ToString())

                    Next

                    ReturnList = OrdYearList

                Case "Rep_OrdinalYearsLess1"

                    Dim OrdYearList As New List(Of String)
                    ' Populate the list with ordinal years from 1 to 40
                    For i As Integer = 2 To 40

                        OrdYearList.Add(i.ToString())

                    Next

                    ReturnList = OrdYearList

                Case "Rep_RealYears"

                    ReturnList = GetListFromNR("Year", ModelID)

                Case "Rep_ExistNew"

                    ReturnList = New List(Of String)({"Existing", "New"})

                Case "Rep_RepayAnnu"

                    ReturnList = New List(Of String)({"Repayment", "Annuity"})

                Case "Rep_FundingRateType"

                    ReturnList = New List(Of String)({"Fixed", "RPI", "Cap", "Variable", "Collar", "Index"})

                Case "Rep_FundingCompounding"

                    ReturnList = New List(Of String)({"Compounding", "Flat rate"})

                Case "Rep_FundingAmortisation"

                    ReturnList = New List(Of String)({"Straight-line", "Manual"})

                Case "Rep_YesNo"

                    ReturnList = New List(Of String)({"Yes", "No"})

                Case "Rep_NIFFixed"

                    ReturnList = New List(Of String)({"NIF basis", "Fixed amount"})

                Case "Rep_TanIntang"

                    ReturnList = New List(Of String)({"Tangible", "Intangible"})

                Case "Rep_ProfileMod"

                    ReturnList = New List(Of String)({"S-curve", "Profile", "Straight line"})

                Case "Rep_GrantMod"

                    ReturnList = New List(Of String)({"Input", "Straight line", "Profile"})

                Case "Rep_SHGCalcBas"

                    ReturnList = New List(Of String)({"Per Unit", "Total SHG", "% of Scheme costs"})

                Case "Rep_DevCostMod"

                    ReturnList = New List(Of String)({"Total Scheme Cost", "Scheme Costs per unit", "Land and Build cost per unit", "Total Land and Build costs"})

                Case "Rep_DepnMode"

                    ReturnList = New List(Of String)({"Diminishing Value", "Straightline"})

                Case "Rep_ImpCharge"

                    ReturnList = New List(Of String)({"Impairment", "Dec in Valuation of Hsg Prop"})

                Case "Rep_OnCostCalcMeth"

                    ReturnList = New List(Of String)({"Total", "Percentage"})

                Case "Rep_CommUncomm"

                    ReturnList = New List(Of String)({"Uncommitted", "Committed"})

                Case "Rep_UnToManProfile"

                    ReturnList = New List(Of String)({"Unit into Mgmt", "Profile"})

                Case "Rep_CPIRPI"

                    ReturnList = New List(Of String)({"CPI", "RPI"})

                Case "Rep_StaOthCust"

                    ReturnList = New List(Of String)({"Staff Costs", "Other Costs"})

                Case "Rep_PERINP"

                    ReturnList = New List(Of String)({"Percentage", "Input"})

                Case "Rep_RecyRepy"

                    ReturnList = New List(Of String)({"Recycled", "Repaid"})

                Case "Rep_FOASales"

                    ReturnList = New List(Of String)({"No", "Current NBV", "OtherPrice"})

                Case Else

                    ReturnList = GetListFromNR(Right(ID, Len(ID) - 4), ModelID)

            End Select

            If InsertBlankFirstItem Then ReturnList.Insert(0, "<blank>")

            EditorReturn.ListItems = ReturnList
            EditorReturn.RetCombo.Items.AddRange(EditorReturn.ListItems)


            CType(DefaultCombo, System.ComponentModel.ISupportInitialize).EndInit()

            'If Not Repositaries(ID).IsInitialised Then

            '    Repositaries(ID).Initialise()

            'End If

            Return EditorReturn



        End Function

        Public Shared Function GetNumericDropEditor(intStart As Integer, intEnd As Integer, Optional ByVal intStep As Integer = 1) As AbovoRespositaryItem

            Dim EditorReturn As New AbovoRespositaryItem
            Dim DefaultCombo As New RepositoryItemComboBox

            CType(DefaultCombo, System.ComponentModel.ISupportInitialize).BeginInit()

            With DefaultCombo
                .Appearance.Font = DefaultFont
                .Appearance.Options.UseFont = True
                .AppearanceDropDown.Font = DefaultFont
                .AppearanceDropDown.Options.UseFont = True
                .AutoHeight = False
                .TextEditStyle = TextEditStyles.DisableTextEditor
                .BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
                .Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
                .LookAndFeel.UseDefaultLookAndFeel = True
            End With

            Dim x As Integer = intStart
            Dim OutList As New List(Of String)

            For x = intStart To intEnd Step intStep

                OutList.Add(x.ToString)

            Next

            EditorReturn.RepType = "CMB"
            EditorReturn.RetCombo = DefaultCombo
            EditorReturn.ListItems = OutList
            EditorReturn.RetCombo.Items.AddRange(EditorReturn.ListItems)

            Return EditorReturn

        End Function
        Public Shared Function GetList(ID As String, ModelID As Integer) As List(Of String)

            Dim ListReturn As List(Of String)


            If Len(ID) > 7 Then

                If Left(ID, 7) = "Rep_NOS" Then

                    Dim NumList As New List(Of String)

                    Dim BreakPos As Integer = InStrRev(ID, "#")

                    If BreakPos = 0 Then GoTo SkipNOS

                    ' Populate the list with ordinal years from 1 to 40

                    Dim EndNo As Integer = CInt(Mid(ID, BreakPos + 1))

                    Dim StartNo As Integer = CInt(Mid(ID, 8, BreakPos - 8))


                    For i As Integer = StartNo To EndNo

                        NumList.Add(i.ToString())

                    Next

                    ListReturn = NumList

                    Return ListReturn

                    Exit Function

                End If

            End If

SkipNOS:


            Select Case ID

                Case "Rep_OwnMan"


                    ListReturn = New List(Of String)({"Owned", "Managed"})

                Case "Rep_ExistNew"

                    ListReturn = New List(Of String)({"Existing", "New"})

                Case "Rep_YEByP"

                    ListReturn = New List(Of String)({"Year End", "By Period"})

                Case "Rep_RepayAnnu"

                    ListReturn = New List(Of String)({"Repayment", "Annuity"})

                Case "Rep_FundingRateType"

                    ListReturn = New List(Of String)({"Fixed", "RPI", "Cap", "Variable", "Collar", "Index"})

                Case "Rep_FundingCompounding"

                    ListReturn = New List(Of String)({"Compounding", "Flat rate"})

                Case "Rep_FundingAmortisation"

                    ListReturn = New List(Of String)({"Straight-line", "Manual"})

                Case "Rep_CMCats"


                    ListReturn = GetListFromNR("CMCats", ModelID)

                Case "Rep_RecyRepy"

                    ListReturn = New List(Of String)({"Recycled", "Repaid"})

                Case "Rep_ValueBasis"


                    ListReturn = New List(Of String)({"EUV-SH", "Cost", "OMV", "MVT"})


                Case "Rep_IncExclude"

                    ListReturn = New List(Of String)({"Include", "Exclude"})

                Case "Rep_FFRRentType"

                    ListReturn = GetListFromNR("FFRRentType", ModelID)


                Case "Rep_OrdinalYears"

                    Dim OrdYearList As New List(Of String)
                    ' Populate the list with ordinal years from 1 to 40
                    For i As Integer = 1 To 40

                        OrdYearList.Add(i.ToString())

                    Next

                    ListReturn = OrdYearList

                Case "Rep_OrdinalYearsLess1"

                    Dim OrdYearList As New List(Of String)
                    ' Populate the list with ordinal years from 1 to 40
                    For i As Integer = 2 To 40

                        OrdYearList.Add(i.ToString())

                    Next

                    ListReturn = OrdYearList

                Case "Rep_RealYears"

                    ListReturn = GetListFromNR("Year", ModelID)

                Case "Rep_YesNo"

                    ListReturn = New List(Of String)({"Yes", "No"})

                Case "Rep_TanIntang"


                    ListReturn = New List(Of String)({"Tangible", "Intangible"})

                Case "Rep_ProfileMod"

                    ListReturn = New List(Of String)({"S-curve", "Profile", "Straight line"})

                Case "Rep_GrantMod"

                    ListReturn = New List(Of String)({"Input", "Straight line", "Profile"})

                Case "Rep_SHGCalcBas"


                    ListReturn = New List(Of String)({"Per Unit", "Total SHG", "% of Scheme costs"})

                Case "Rep_DevCostMod"

                    ListReturn = New List(Of String)({"Total Scheme Cost", "Scheme Costs per unit", "Land and Build cost per unit", "Total Land and Build costs"})

                Case "Rep_DepnMode"

                    ListReturn = New List(Of String)({"Diminishing Value", "Straightline"})

                Case "Rep_ImpCharge"

                    ListReturn = New List(Of String)({"Impairment", "Dec in Valuation of Hsg Prop"})

                Case "Rep_OnCostCalcMeth"

                    ListReturn = New List(Of String)({"Total", "Percentage"})

                Case "Rep_CommUncomm"

                    ListReturn = New List(Of String)({"Uncommitted", "Committed"})

                Case "Rep_UnToManProfile"

                    ListReturn = New List(Of String)({"Unit into Mgmt", "Profile"})

                Case "Rep_CPIRPI"

                    ListReturn = New List(Of String)({"CPI", "RPI"})

                Case "Rep_StaOthCust"

                    ListReturn = New List(Of String)({"Staff Costs", "Other Costs"})
                Case Else

                    ListReturn = GetListFromNR(Right(ID, Len(ID) - 4), ModelID)

            End Select


            Return ListReturn

        End Function

        Public Shared Function GetListFromNR(RangeName As String, SetModelID As Integer) As List(Of String)

            Dim NRItems As New List(Of String)
            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = ExcelModels(SetModelID).WB.DefinedNames.GetDefinedName(RangeName).Range
            Dim clCell As DevExpress.Spreadsheet.Cell = CRTargetRange(0, 0)

            If CRTargetRange.RowCount > CRTargetRange.ColumnCount Then

                For x = 0 To CRTargetRange.RowCount - 1

                    clCell = CRTargetRange(x, 0)
                    If Len(clCell.DisplayText) > 0 Then NRItems.Add(clCell.DisplayText)

                Next

            Else

                For x = 0 To CRTargetRange.ColumnCount - 1

                    clCell = CRTargetRange(0, x)
                    If Len(clCell.DisplayText) > 0 Then NRItems.Add(clCell.DisplayText)

                Next

            End If

            Return NRItems

        End Function
        Public Class AbovoRespositaryItem

            Public ListItems As List(Of String)
            Public IsInitialised As Boolean
            Public Index As Integer
            Public RepType As String
            Public RetCombo As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
            Public Sub Initialise()
            End Sub
        End Class
        Public Function GetRepositaryItem(SoughtRepItem As String) As AbovoRespositaryItem
            Dim RetARI As New AbovoRespositaryItem
            Select Case SoughtRepItem
                Case "CmbOwnedManaged"
                    RetARI.RetCombo = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
            End Select
            Return RetARI
        End Function
        Class PercentSpinEdit
            Public Shared Function GetControl() As RepositoryItemSpinEdit
                Dim SE As New RepositoryItemSpinEdit

                Return SE
                SE.Appearance.BackColor = System.Drawing.Color.White
                SE.Appearance.BorderColor = System.Drawing.Color.White
                SE.Appearance.Options.UseBackColor = True
                SE.Appearance.Options.UseBorderColor = True
                SE.AutoHeight = False
                SE.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
                SE.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
                SE.ButtonsStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
                SE.Increment = 0.000025
                SE.LookAndFeel.UseDefaultLookAndFeel = False
                SE.MaskSettings.Set("mask", "p")
                SE.MaxValue = 100
                'SE.Name = "RepositoryItemSpinEditInitialRateNewLettings"
                SE.UseMaskAsDisplayFormat = True
                CType(SE, System.ComponentModel.ISupportInitialize).BeginInit()
                SE.IsFloatValue = False
                Return SE




            End Function

        End Class


        Public Class OMType
            Public Property OMName() As String
            Public Shared Function Init() As List(Of OMType)
                Return New List(Of OMType)() From {
                     New OMType() With {.OMName = "Owned"},
                     New OMType() With {.OMName = "Managed"}
                   }
            End Function
        End Class
        Public Class SOCIStock
            <Display(Order:=-1)>
            Public Property SOCIStockID() As Integer
            Public Property CategoryID() As Integer
            Public Property SOCIStockName() As String
            Public Shared Function Init() As List(Of SOCIStock)
                Return New List(Of SOCIStock)() From {
                        New SOCIStock() With {.SOCIStockID = 0, .CategoryID = 1, .SOCIStockName = "Gen Needs"},
                        New SOCIStock() With {.SOCIStockID = 1, .CategoryID = 0, .SOCIStockName = "LCHO"},
                        New SOCIStock() With {.SOCIStockID = 2, .CategoryID = 1, .SOCIStockName = "Supported"},
                        New SOCIStock() With {.SOCIStockID = 3, .CategoryID = 0, .SOCIStockName = "Other"},
                        New SOCIStock() With {.SOCIStockID = 4, .CategoryID = 0, .SOCIStockName = "N/A"},
                        New SOCIStock() With {.SOCIStockID = 5, .CategoryID = 0, .SOCIStockName = "Supported"},
                        New SOCIStock() With {.SOCIStockID = 5, .CategoryID = 0, .SOCIStockName = "Non-social"}
                    }
            End Function
            Public Function GetSOCICategoryByName(ByVal MyStockName As String) As Integer

                Dim selectedValue As SOCIStock
                selectedValue = Init.Find(Function(p) p.SOCIStockName = MyStockName)
                Return selectedValue.CategoryID

            End Function

        End Class
        Public Class SOCIRentType
            <Display(Order:=-1)>
            Public Property SOCIRentTypetID() As Integer
            Public Property SOCIRentName() As String
            <Display(Order:=-1)>
            Public Property CategoryID() As Integer


            Public Shared Function Init() As List(Of SOCIRentType)
                Return New List(Of SOCIRentType)() From {
                        New SOCIRentType() With {.SOCIRentTypetID = 0, .SOCIRentName = "N/A", .CategoryID = 0},
                        New SOCIRentType() With {.SOCIRentTypetID = 1, .SOCIRentName = "Social Rent", .CategoryID = 1},
                        New SOCIRentType() With {.SOCIRentTypetID = 2, .SOCIRentName = "Aff Rent", .CategoryID = 1}
                    }
            End Function
            Public Shared Function GetSOCIRentTypeByCategory(ByVal categoryId As Integer) As List(Of SOCIRentType)

                Return Init().Where(Function(p) p.CategoryID = categoryId).ToList()

            End Function

        End Class
    End Class

End Namespace
