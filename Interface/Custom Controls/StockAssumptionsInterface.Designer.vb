Imports DevExpress.Data
Imports Abovo.AbovoAppCls
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class StockAssumptionsInterface
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim UnboundSourceProperty1 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty2 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty3 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty4 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty5 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty6 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty7 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty8 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty9 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty10 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty11 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty12 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty13 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim UnboundSourceProperty14 As DevExpress.Data.UnboundSourceProperty = New DevExpress.Data.UnboundSourceProperty()
        Dim EditorButtonImageOptions1 As DevExpress.XtraEditors.Controls.EditorButtonImageOptions = New DevExpress.XtraEditors.Controls.EditorButtonImageOptions()
        Dim SerializableAppearanceObject1 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim SerializableAppearanceObject2 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim SerializableAppearanceObject3 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim SerializableAppearanceObject4 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim WindowsUIButtonImageOptions1 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(StockAssumptionsInterface))
        Dim WindowsUIButtonImageOptions2 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions3 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions4 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions5 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions6 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions7 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Me.UnboundSourceStocks = New DevExpress.Data.UnboundSource(Me.components)
        Me.GridControlStockGrid = New DevExpress.XtraGrid.GridControl()
        Me.GridViewStockNumbers = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.colPropertyStockDescription1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colPropertyOwnedManaged1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.RepositoryItemComboBoxOwnedManaged = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.colPropertySOCIStockType1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.RepositoryItemLookUpEditSOCIStockType = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.colPropertySOCIRentType1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.RepositoryItemLookUpEditSOCIRentType = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.colPropertyCurrentStockNumbers1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.RepositoryItemIntegerEdit = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colPropertyPreBPlanStartDateNewBuild1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colPropertyPreBPlanStartDateDemolitions1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colPropertyPreBPlanStartDateRTBs1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colPropertyPreBPlanStartDateOtherDisposals1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colPropertyExistingStocksCalc1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colPropertyNewLettings1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colPropertyInitialRateNewLettings1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.RepositoryItemSpinEditInitialRateNewLettings = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.colPropertyTotalOpeningStockCalc1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.RepositoryItemComboBoxSOCIRent = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.RepositoryItemComboBoxSOCIStockType = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.RepositoryItemTextEditThousandSep = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.CardView1 = New DevExpress.XtraGrid.Views.Card.CardView()
        Me.LayoutViewStock = New DevExpress.XtraGrid.Views.Layout.LayoutView()
        Me.LayoutViewColumn1 = New DevExpress.XtraGrid.Columns.LayoutViewColumn()
        Me.layoutViewField_LayoutViewColumn1 = New DevExpress.XtraGrid.Views.Layout.LayoutViewField()
        Me.LayoutViewCard1 = New DevExpress.XtraGrid.Views.Layout.LayoutViewCard()
        Me.TablePanelStockAssumptions = New DevExpress.Utils.Layout.TablePanel()
        Me.WindowsUIButtonPanelItemEdit = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        CType(Me.UnboundSourceStocks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridControlStockGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridViewStockNumbers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemComboBoxOwnedManaged, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemLookUpEditSOCIStockType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemLookUpEditSOCIRentType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemIntegerEdit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEditInitialRateNewLettings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemComboBoxSOCIRent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemComboBoxSOCIStockType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEditThousandSep, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CardView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutViewStock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.layoutViewField_LayoutViewColumn1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutViewCard1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TablePanelStockAssumptions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanelStockAssumptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'UnboundSourceStocks
        '
        UnboundSourceProperty1.DisplayName = "Stock Number"
        UnboundSourceProperty1.Name = "IntIdentifier"
        UnboundSourceProperty1.PropertyType = GetType(Integer)
        UnboundSourceProperty2.DisplayName = "Description"
        UnboundSourceProperty2.Name = "PropertyStockDescription"
        UnboundSourceProperty2.PropertyType = GetType(String)
        UnboundSourceProperty3.DisplayName = "Owned or Managed"
        UnboundSourceProperty3.Name = "PropertyOwnedManaged"
        UnboundSourceProperty3.PropertyType = GetType(String)
        UnboundSourceProperty4.DisplayName = "SOCI Stock Type Description"
        UnboundSourceProperty4.Name = "PropertySOCIStockType"
        UnboundSourceProperty4.PropertyType = GetType(String)
        UnboundSourceProperty5.DisplayName = "SOCI Rent Type"
        UnboundSourceProperty5.Name = "PropertySOCIRentType"
        UnboundSourceProperty5.PropertyType = GetType(String)
        UnboundSourceProperty6.DisplayName = "Current Stock Numbers"
        UnboundSourceProperty6.Name = "PropertyCurrentStockNumbers"
        UnboundSourceProperty6.PropertyType = GetType(Integer)
        UnboundSourceProperty7.DisplayName = "Pre BPlan Start Date New Build"
        UnboundSourceProperty7.Name = "PropertyPreBPlanStartDateNewBuild"
        UnboundSourceProperty7.PropertyType = GetType(Integer)
        UnboundSourceProperty8.DisplayName = "Pre BPlan Start Date Demolitions"
        UnboundSourceProperty8.Name = "PropertyPreBPlanStartDateDemolitions"
        UnboundSourceProperty8.PropertyType = GetType(Integer)
        UnboundSourceProperty9.DisplayName = "Pre BPlan Start Date RTBs"
        UnboundSourceProperty9.Name = "PropertyPreBPlanStartDateRTBs"
        UnboundSourceProperty9.PropertyType = GetType(Integer)
        UnboundSourceProperty10.DisplayName = "Pre BPlan Start Date Other Disposals"
        UnboundSourceProperty10.Name = "PropertyPreBPlanStartDateOtherDisposals"
        UnboundSourceProperty10.PropertyType = GetType(Integer)
        UnboundSourceProperty11.DisplayName = "Existing Stock"
        UnboundSourceProperty11.Name = "PropertyExistingStocksCalc"
        UnboundSourceProperty11.PropertyType = GetType(Integer)
        UnboundSourceProperty12.DisplayName = "New Lettings"
        UnboundSourceProperty12.Name = "PropertyNewLettings"
        UnboundSourceProperty12.PropertyType = GetType(Integer)
        UnboundSourceProperty13.DisplayName = "Initial Rate of New Lettings"
        UnboundSourceProperty13.Name = "PropertyInitialRateNewLettings"
        UnboundSourceProperty13.PropertyType = GetType(Single)
        UnboundSourceProperty14.DisplayName = "Total Opening Stock"
        UnboundSourceProperty14.Name = "PropertyTotalOpeningStockCalc"
        UnboundSourceProperty14.PropertyType = GetType(Integer)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty1)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty2)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty3)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty4)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty5)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty6)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty7)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty8)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty9)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty10)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty11)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty12)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty13)
        Me.UnboundSourceStocks.Properties.Add(UnboundSourceProperty14)
        '
        'GridControlStockGrid
        '
        Me.TablePanelStockAssumptions.SetColumn(Me.GridControlStockGrid, 0)
        Me.GridControlStockGrid.DataSource = Me.UnboundSourceStocks
        Me.GridControlStockGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridControlStockGrid.EmbeddedNavigator.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GridControlStockGrid.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GridControlStockGrid.Location = New System.Drawing.Point(43, 105)
        Me.GridControlStockGrid.MainView = Me.GridViewStockNumbers
        Me.GridControlStockGrid.Margin = New System.Windows.Forms.Padding(24, 3, 24, 3)
        Me.GridControlStockGrid.Name = "GridControlStockGrid"
        Me.GridControlStockGrid.Padding = New System.Windows.Forms.Padding(24, 0, 24, 0)
        Me.GridControlStockGrid.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemComboBoxOwnedManaged, Me.RepositoryItemComboBoxSOCIRent, Me.RepositoryItemSpinEditInitialRateNewLettings, Me.RepositoryItemComboBoxSOCIStockType, Me.RepositoryItemIntegerEdit, Me.RepositoryItemLookUpEditSOCIRentType, Me.RepositoryItemLookUpEditSOCIStockType, Me.RepositoryItemTextEditThousandSep})
        Me.TablePanelStockAssumptions.SetRow(Me.GridControlStockGrid, 1)
        Me.GridControlStockGrid.Size = New System.Drawing.Size(1483, 755)
        Me.GridControlStockGrid.TabIndex = 0
        Me.GridControlStockGrid.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridViewStockNumbers, Me.CardView1, Me.LayoutViewStock})
        '
        'GridViewStockNumbers
        '
        Me.GridViewStockNumbers.Appearance.Empty.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.Empty.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.Empty.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.Empty.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.EvenRow.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.EvenRow.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.EvenRow.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GridViewStockNumbers.Appearance.EvenRow.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.EvenRow.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.EvenRow.Options.UseFont = True
        Me.GridViewStockNumbers.Appearance.FixedLine.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.FixedLine.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.FixedLine.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.FixedLine.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.GridViewStockNumbers.Appearance.FocusedRow.BorderColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.GridViewStockNumbers.Appearance.FocusedRow.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GridViewStockNumbers.Appearance.FocusedRow.ForeColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.FocusedRow.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.FocusedRow.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.FocusedRow.Options.UseFont = True
        Me.GridViewStockNumbers.Appearance.FocusedRow.Options.UseForeColor = True
        Me.GridViewStockNumbers.Appearance.HeaderPanel.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HeaderPanel.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HeaderPanel.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GridViewStockNumbers.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.GridViewStockNumbers.Appearance.HeaderPanel.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.HeaderPanel.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.HeaderPanel.Options.UseFont = True
        Me.GridViewStockNumbers.Appearance.HeaderPanel.Options.UseForeColor = True
        Me.GridViewStockNumbers.Appearance.HeaderPanel.Options.UseTextOptions = True
        Me.GridViewStockNumbers.Appearance.HeaderPanel.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.GridViewStockNumbers.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.GridViewStockNumbers.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HideSelectionRow.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.HorzLine.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.ForeColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseForeColor = True
        Me.GridViewStockNumbers.Appearance.OddRow.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.OddRow.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.OddRow.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GridViewStockNumbers.Appearance.OddRow.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.OddRow.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.OddRow.Options.UseFont = True
        Me.GridViewStockNumbers.Appearance.VertLine.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.VertLine.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.VertLine.ForeColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.VertLine.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.VertLine.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.VertLine.Options.UseForeColor = True
        Me.GridViewStockNumbers.Appearance.ViewCaption.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GridViewStockNumbers.Appearance.ViewCaption.Options.UseFont = True
        Me.GridViewStockNumbers.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.GridViewStockNumbers.ColumnPanelRowHeight = 148
        Me.GridViewStockNumbers.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.colPropertyStockDescription1, Me.colPropertyOwnedManaged1, Me.colPropertySOCIStockType1, Me.colPropertySOCIRentType1, Me.colPropertyCurrentStockNumbers1, Me.colPropertyPreBPlanStartDateNewBuild1, Me.colPropertyPreBPlanStartDateDemolitions1, Me.colPropertyPreBPlanStartDateRTBs1, Me.colPropertyPreBPlanStartDateOtherDisposals1, Me.colPropertyExistingStocksCalc1, Me.colPropertyNewLettings1, Me.colPropertyInitialRateNewLettings1, Me.colPropertyTotalOpeningStockCalc1})
        Me.GridViewStockNumbers.DetailHeight = 399
        Me.GridViewStockNumbers.FixedLineWidth = 1
        Me.GridViewStockNumbers.GridControl = Me.GridControlStockGrid
        Me.GridViewStockNumbers.Name = "GridViewStockNumbers"
        Me.GridViewStockNumbers.OptionsCustomization.AllowColumnMoving = False
        Me.GridViewStockNumbers.OptionsCustomization.AllowFilter = False
        Me.GridViewStockNumbers.OptionsCustomization.AllowGroup = False
        Me.GridViewStockNumbers.OptionsCustomization.AllowRowSizing = True
        Me.GridViewStockNumbers.OptionsCustomization.AllowSort = False
        Me.GridViewStockNumbers.OptionsEditForm.PopupEditFormWidth = 971
        Me.GridViewStockNumbers.OptionsMenu.EnableColumnMenu = False
        Me.GridViewStockNumbers.OptionsSelection.EnableAppearanceFocusedCell = False
        Me.GridViewStockNumbers.OptionsSelection.EnableAppearanceFocusedRow = False
        Me.GridViewStockNumbers.OptionsView.AllowHtmlDrawHeaders = True
        Me.GridViewStockNumbers.OptionsView.ColumnAutoWidth = False
        Me.GridViewStockNumbers.OptionsView.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowForFocusedRow
        Me.GridViewStockNumbers.OptionsView.ShowGroupPanel = False
        Me.GridViewStockNumbers.OptionsView.ShowIndicator = False
        Me.GridViewStockNumbers.RowHeight = 23
        Me.GridViewStockNumbers.ScrollStyle = DevExpress.XtraGrid.Views.Grid.ScrollStyleFlags.LiveVertScroll
        '
        'colPropertyStockDescription1
        '
        Me.colPropertyStockDescription1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyStockDescription1.AppearanceCell.Options.UseFont = True
        Me.colPropertyStockDescription1.AppearanceCell.Options.UseTextOptions = True
        Me.colPropertyStockDescription1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colPropertyStockDescription1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyStockDescription1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyStockDescription1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyStockDescription1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyStockDescription1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyStockDescription1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyStockDescription1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyStockDescription1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colPropertyStockDescription1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyStockDescription1.FieldName = "PropertyStockDescription"
        Me.colPropertyStockDescription1.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        Me.colPropertyStockDescription1.MinWidth = 42
        Me.colPropertyStockDescription1.Name = "colPropertyStockDescription1"
        Me.colPropertyStockDescription1.Visible = True
        Me.colPropertyStockDescription1.VisibleIndex = 0
        Me.colPropertyStockDescription1.Width = 277
        '
        'colPropertyOwnedManaged1
        '
        Me.colPropertyOwnedManaged1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyOwnedManaged1.AppearanceCell.Options.UseFont = True
        Me.colPropertyOwnedManaged1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyOwnedManaged1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyOwnedManaged1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyOwnedManaged1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyOwnedManaged1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyOwnedManaged1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyOwnedManaged1.Caption = "Owned / Managed"
        Me.colPropertyOwnedManaged1.ColumnEdit = Me.RepositoryItemComboBoxOwnedManaged
        Me.colPropertyOwnedManaged1.FieldName = "PropertyOwnedManaged"
        Me.colPropertyOwnedManaged1.MinWidth = 42
        Me.colPropertyOwnedManaged1.Name = "colPropertyOwnedManaged1"
        Me.colPropertyOwnedManaged1.Visible = True
        Me.colPropertyOwnedManaged1.VisibleIndex = 1
        Me.colPropertyOwnedManaged1.Width = 195
        '
        'RepositoryItemComboBoxOwnedManaged
        '
        Me.RepositoryItemComboBoxOwnedManaged.Appearance.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!)
        Me.RepositoryItemComboBoxOwnedManaged.Appearance.Options.UseFont = True
        Me.RepositoryItemComboBoxOwnedManaged.AppearanceDropDown.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemComboBoxOwnedManaged.AppearanceDropDown.Options.UseFont = True
        Me.RepositoryItemComboBoxOwnedManaged.AutoHeight = False
        Me.RepositoryItemComboBoxOwnedManaged.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple
        Me.RepositoryItemComboBoxOwnedManaged.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.RepositoryItemComboBoxOwnedManaged.Items.AddRange(New Object() {"Owned", "Managed"})
        Me.RepositoryItemComboBoxOwnedManaged.LookAndFeel.UseDefaultLookAndFeel = False
        Me.RepositoryItemComboBoxOwnedManaged.Name = "RepositoryItemComboBoxOwnedManaged"
        '
        'colPropertySOCIStockType1
        '
        Me.colPropertySOCIStockType1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertySOCIStockType1.AppearanceCell.Options.UseFont = True
        Me.colPropertySOCIStockType1.AppearanceCell.Options.UseTextOptions = True
        Me.colPropertySOCIStockType1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colPropertySOCIStockType1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertySOCIStockType1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertySOCIStockType1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertySOCIStockType1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertySOCIStockType1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertySOCIStockType1.AppearanceHeader.Options.UseFont = True
        Me.colPropertySOCIStockType1.Caption = "SOCI Stock Type"
        Me.colPropertySOCIStockType1.ColumnEdit = Me.RepositoryItemLookUpEditSOCIStockType
        Me.colPropertySOCIStockType1.FieldName = "PropertySOCIStockType"
        Me.colPropertySOCIStockType1.MinWidth = 42
        Me.colPropertySOCIStockType1.Name = "colPropertySOCIStockType1"
        Me.colPropertySOCIStockType1.Visible = True
        Me.colPropertySOCIStockType1.VisibleIndex = 2
        Me.colPropertySOCIStockType1.Width = 232
        '
        'RepositoryItemLookUpEditSOCIStockType
        '
        Me.RepositoryItemLookUpEditSOCIStockType.Appearance.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemLookUpEditSOCIStockType.Appearance.Options.UseFont = True
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDown.BackColor = System.Drawing.Color.White
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDown.BorderColor = System.Drawing.Color.White
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDown.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDown.Options.UseBackColor = True
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDown.Options.UseBorderColor = True
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDown.Options.UseFont = True
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDownHeader.BackColor = System.Drawing.Color.White
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDownHeader.BorderColor = System.Drawing.Color.White
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDownHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDownHeader.Options.UseBackColor = True
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDownHeader.Options.UseBorderColor = True
        Me.RepositoryItemLookUpEditSOCIStockType.AppearanceDropDownHeader.Options.UseFont = True
        Me.RepositoryItemLookUpEditSOCIStockType.AutoHeight = False
        Me.RepositoryItemLookUpEditSOCIStockType.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        SerializableAppearanceObject1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        SerializableAppearanceObject1.Options.UseForeColor = True
        Me.RepositoryItemLookUpEditSOCIStockType.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo, "", -1, True, True, False, EditorButtonImageOptions1, New DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), SerializableAppearanceObject1, SerializableAppearanceObject2, SerializableAppearanceObject3, SerializableAppearanceObject4, "", Nothing, Nothing, DevExpress.Utils.ToolTipAnchor.[Default])})
        Me.RepositoryItemLookUpEditSOCIStockType.ButtonsStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.RepositoryItemLookUpEditSOCIStockType.LookAndFeel.UseDefaultLookAndFeel = False
        Me.RepositoryItemLookUpEditSOCIStockType.Name = "RepositoryItemLookUpEditSOCIStockType"
        Me.RepositoryItemLookUpEditSOCIStockType.Padding = New System.Windows.Forms.Padding(5)
        Me.RepositoryItemLookUpEditSOCIStockType.PopupBorderStyle = DevExpress.XtraEditors.Controls.PopupBorderStyles.NoBorder
        '
        'colPropertySOCIRentType1
        '
        Me.colPropertySOCIRentType1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertySOCIRentType1.AppearanceCell.Options.UseFont = True
        Me.colPropertySOCIRentType1.AppearanceCell.Options.UseTextOptions = True
        Me.colPropertySOCIRentType1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colPropertySOCIRentType1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertySOCIRentType1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertySOCIRentType1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertySOCIRentType1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertySOCIRentType1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertySOCIRentType1.AppearanceHeader.Options.UseFont = True
        Me.colPropertySOCIRentType1.Caption = "SOCI Rent Class"
        Me.colPropertySOCIRentType1.ColumnEdit = Me.RepositoryItemLookUpEditSOCIRentType
        Me.colPropertySOCIRentType1.FieldName = "PropertySOCIRentType"
        Me.colPropertySOCIRentType1.MinWidth = 42
        Me.colPropertySOCIRentType1.Name = "colPropertySOCIRentType1"
        Me.colPropertySOCIRentType1.Visible = True
        Me.colPropertySOCIRentType1.VisibleIndex = 3
        Me.colPropertySOCIRentType1.Width = 212
        '
        'RepositoryItemLookUpEditSOCIRentType
        '
        Me.RepositoryItemLookUpEditSOCIRentType.Appearance.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemLookUpEditSOCIRentType.Appearance.Options.UseFont = True
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDown.BackColor = System.Drawing.Color.White
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDown.BorderColor = System.Drawing.Color.White
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDown.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDown.Options.UseBackColor = True
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDown.Options.UseBorderColor = True
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDown.Options.UseFont = True
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDownHeader.BackColor = System.Drawing.Color.White
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDownHeader.BorderColor = System.Drawing.Color.White
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDownHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDownHeader.Options.UseBackColor = True
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDownHeader.Options.UseBorderColor = True
        Me.RepositoryItemLookUpEditSOCIRentType.AppearanceDropDownHeader.Options.UseFont = True
        Me.RepositoryItemLookUpEditSOCIRentType.AutoHeight = False
        Me.RepositoryItemLookUpEditSOCIRentType.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.RepositoryItemLookUpEditSOCIRentType.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.RepositoryItemLookUpEditSOCIRentType.ButtonsStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.RepositoryItemLookUpEditSOCIRentType.LookAndFeel.UseDefaultLookAndFeel = False
        Me.RepositoryItemLookUpEditSOCIRentType.Name = "RepositoryItemLookUpEditSOCIRentType"
        Me.RepositoryItemLookUpEditSOCIRentType.Padding = New System.Windows.Forms.Padding(5)
        Me.RepositoryItemLookUpEditSOCIRentType.PopupBorderStyle = DevExpress.XtraEditors.Controls.PopupBorderStyles.NoBorder
        '
        'colPropertyCurrentStockNumbers1
        '
        Me.colPropertyCurrentStockNumbers1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyCurrentStockNumbers1.AppearanceCell.Options.UseFont = True
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyCurrentStockNumbers1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyCurrentStockNumbers1.Caption = "Current Stock"
        Me.colPropertyCurrentStockNumbers1.ColumnEdit = Me.RepositoryItemIntegerEdit
        Me.colPropertyCurrentStockNumbers1.DisplayFormat.FormatString = "N"
        Me.colPropertyCurrentStockNumbers1.FieldName = "PropertyCurrentStockNumbers"
        Me.colPropertyCurrentStockNumbers1.MinWidth = 42
        Me.colPropertyCurrentStockNumbers1.Name = "colPropertyCurrentStockNumbers1"
        Me.colPropertyCurrentStockNumbers1.Visible = True
        Me.colPropertyCurrentStockNumbers1.VisibleIndex = 4
        Me.colPropertyCurrentStockNumbers1.Width = 159
        '
        'RepositoryItemIntegerEdit
        '
        Me.RepositoryItemIntegerEdit.Appearance.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemIntegerEdit.Appearance.Options.UseFont = True
        Me.RepositoryItemIntegerEdit.AutoHeight = False
        Me.RepositoryItemIntegerEdit.DisplayFormat.FormatString = "#,###,##0"
        Me.RepositoryItemIntegerEdit.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.RepositoryItemIntegerEdit.EditFormat.FormatString = "#,###,##0"
        Me.RepositoryItemIntegerEdit.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.RepositoryItemIntegerEdit.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
        Me.RepositoryItemIntegerEdit.MaskSettings.Set("MaskManagerSignature", "allowNull=False")
        Me.RepositoryItemIntegerEdit.MaskSettings.Set("mask", "D")
        Me.RepositoryItemIntegerEdit.Name = "RepositoryItemIntegerEdit"
        '
        'colPropertyPreBPlanStartDateNewBuild1
        '
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceCell.Options.UseFont = True
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyPreBPlanStartDateNewBuild1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyPreBPlanStartDateNewBuild1.Caption = "Pre Plan New Build"
        Me.colPropertyPreBPlanStartDateNewBuild1.ColumnEdit = Me.RepositoryItemIntegerEdit
        Me.colPropertyPreBPlanStartDateNewBuild1.DisplayFormat.FormatString = "n2"
        Me.colPropertyPreBPlanStartDateNewBuild1.FieldName = "PropertyPreBPlanStartDateNewBuild"
        Me.colPropertyPreBPlanStartDateNewBuild1.MinWidth = 42
        Me.colPropertyPreBPlanStartDateNewBuild1.Name = "colPropertyPreBPlanStartDateNewBuild1"
        Me.colPropertyPreBPlanStartDateNewBuild1.Visible = True
        Me.colPropertyPreBPlanStartDateNewBuild1.VisibleIndex = 5
        Me.colPropertyPreBPlanStartDateNewBuild1.Width = 159
        '
        'colPropertyPreBPlanStartDateDemolitions1
        '
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceCell.Options.UseFont = True
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyPreBPlanStartDateDemolitions1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyPreBPlanStartDateDemolitions1.Caption = "Pre Plan Demolitions"
        Me.colPropertyPreBPlanStartDateDemolitions1.ColumnEdit = Me.RepositoryItemIntegerEdit
        Me.colPropertyPreBPlanStartDateDemolitions1.FieldName = "PropertyPreBPlanStartDateDemolitions"
        Me.colPropertyPreBPlanStartDateDemolitions1.MinWidth = 42
        Me.colPropertyPreBPlanStartDateDemolitions1.Name = "colPropertyPreBPlanStartDateDemolitions1"
        Me.colPropertyPreBPlanStartDateDemolitions1.Visible = True
        Me.colPropertyPreBPlanStartDateDemolitions1.VisibleIndex = 6
        Me.colPropertyPreBPlanStartDateDemolitions1.Width = 192
        '
        'colPropertyPreBPlanStartDateRTBs1
        '
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceCell.Options.UseFont = True
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyPreBPlanStartDateRTBs1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyPreBPlanStartDateRTBs1.Caption = "Pre Plan RTBs"
        Me.colPropertyPreBPlanStartDateRTBs1.ColumnEdit = Me.RepositoryItemIntegerEdit
        Me.colPropertyPreBPlanStartDateRTBs1.FieldName = "PropertyPreBPlanStartDateRTBs"
        Me.colPropertyPreBPlanStartDateRTBs1.MinWidth = 42
        Me.colPropertyPreBPlanStartDateRTBs1.Name = "colPropertyPreBPlanStartDateRTBs1"
        Me.colPropertyPreBPlanStartDateRTBs1.Visible = True
        Me.colPropertyPreBPlanStartDateRTBs1.VisibleIndex = 7
        Me.colPropertyPreBPlanStartDateRTBs1.Width = 192
        '
        'colPropertyPreBPlanStartDateOtherDisposals1
        '
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceCell.Options.UseFont = True
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyPreBPlanStartDateOtherDisposals1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyPreBPlanStartDateOtherDisposals1.Caption = "Pre Plan Other Disposals"
        Me.colPropertyPreBPlanStartDateOtherDisposals1.ColumnEdit = Me.RepositoryItemIntegerEdit
        Me.colPropertyPreBPlanStartDateOtherDisposals1.DisplayFormat.FormatString = "N"
        Me.colPropertyPreBPlanStartDateOtherDisposals1.FieldName = "PropertyPreBPlanStartDateOtherDisposals"
        Me.colPropertyPreBPlanStartDateOtherDisposals1.MinWidth = 42
        Me.colPropertyPreBPlanStartDateOtherDisposals1.Name = "colPropertyPreBPlanStartDateOtherDisposals1"
        Me.colPropertyPreBPlanStartDateOtherDisposals1.Visible = True
        Me.colPropertyPreBPlanStartDateOtherDisposals1.VisibleIndex = 8
        Me.colPropertyPreBPlanStartDateOtherDisposals1.Width = 192
        '
        'colPropertyExistingStocksCalc1
        '
        Me.colPropertyExistingStocksCalc1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyExistingStocksCalc1.AppearanceCell.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.colPropertyExistingStocksCalc1.AppearanceCell.Options.UseFont = True
        Me.colPropertyExistingStocksCalc1.AppearanceCell.Options.UseForeColor = True
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyExistingStocksCalc1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyExistingStocksCalc1.Caption = "Existing Stocks"
        Me.colPropertyExistingStocksCalc1.DisplayFormat.FormatString = "N"
        Me.colPropertyExistingStocksCalc1.FieldName = "PropertyExistingStocksCalc"
        Me.colPropertyExistingStocksCalc1.MinWidth = 42
        Me.colPropertyExistingStocksCalc1.Name = "colPropertyExistingStocksCalc1"
        Me.colPropertyExistingStocksCalc1.OptionsColumn.AllowEdit = False
        Me.colPropertyExistingStocksCalc1.OptionsColumn.AllowFocus = False
        Me.colPropertyExistingStocksCalc1.OptionsColumn.ReadOnly = True
        Me.colPropertyExistingStocksCalc1.Visible = True
        Me.colPropertyExistingStocksCalc1.VisibleIndex = 9
        Me.colPropertyExistingStocksCalc1.Width = 192
        '
        'colPropertyNewLettings1
        '
        Me.colPropertyNewLettings1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyNewLettings1.AppearanceCell.Options.UseFont = True
        Me.colPropertyNewLettings1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyNewLettings1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyNewLettings1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyNewLettings1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyNewLettings1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyNewLettings1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyNewLettings1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyNewLettings1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyNewLettings1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyNewLettings1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyNewLettings1.Caption = "New Lettings"
        Me.colPropertyNewLettings1.ColumnEdit = Me.RepositoryItemIntegerEdit
        Me.colPropertyNewLettings1.DisplayFormat.FormatString = "N"
        Me.colPropertyNewLettings1.FieldName = "PropertyNewLettings"
        Me.colPropertyNewLettings1.MinWidth = 42
        Me.colPropertyNewLettings1.Name = "colPropertyNewLettings1"
        Me.colPropertyNewLettings1.Visible = True
        Me.colPropertyNewLettings1.VisibleIndex = 10
        Me.colPropertyNewLettings1.Width = 177
        '
        'colPropertyInitialRateNewLettings1
        '
        Me.colPropertyInitialRateNewLettings1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyInitialRateNewLettings1.AppearanceCell.Options.UseFont = True
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyInitialRateNewLettings1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyInitialRateNewLettings1.Caption = "New lettings initial rate"
        Me.colPropertyInitialRateNewLettings1.ColumnEdit = Me.RepositoryItemSpinEditInitialRateNewLettings
        Me.colPropertyInitialRateNewLettings1.FieldName = "PropertyInitialRateNewLettings"
        Me.colPropertyInitialRateNewLettings1.MinWidth = 42
        Me.colPropertyInitialRateNewLettings1.Name = "colPropertyInitialRateNewLettings1"
        Me.colPropertyInitialRateNewLettings1.Visible = True
        Me.colPropertyInitialRateNewLettings1.VisibleIndex = 11
        Me.colPropertyInitialRateNewLettings1.Width = 188
        '
        'RepositoryItemSpinEditInitialRateNewLettings
        '
        Me.RepositoryItemSpinEditInitialRateNewLettings.Appearance.BackColor = System.Drawing.Color.White
        Me.RepositoryItemSpinEditInitialRateNewLettings.Appearance.BorderColor = System.Drawing.Color.White
        Me.RepositoryItemSpinEditInitialRateNewLettings.Appearance.Options.UseBackColor = True
        Me.RepositoryItemSpinEditInitialRateNewLettings.Appearance.Options.UseBorderColor = True
        Me.RepositoryItemSpinEditInitialRateNewLettings.AutoHeight = False
        Me.RepositoryItemSpinEditInitialRateNewLettings.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.RepositoryItemSpinEditInitialRateNewLettings.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.RepositoryItemSpinEditInitialRateNewLettings.ButtonsStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.RepositoryItemSpinEditInitialRateNewLettings.Increment = New Decimal(New Integer() {25, 0, 0, 262144})
        Me.RepositoryItemSpinEditInitialRateNewLettings.LookAndFeel.UseDefaultLookAndFeel = False
        Me.RepositoryItemSpinEditInitialRateNewLettings.MaskSettings.Set("mask", "p")
        Me.RepositoryItemSpinEditInitialRateNewLettings.MaxValue = New Decimal(New Integer() {100, 0, 0, 0})
        Me.RepositoryItemSpinEditInitialRateNewLettings.Name = "RepositoryItemSpinEditInitialRateNewLettings"
        Me.RepositoryItemSpinEditInitialRateNewLettings.UseMaskAsDisplayFormat = True
        '
        'colPropertyTotalOpeningStockCalc1
        '
        Me.colPropertyTotalOpeningStockCalc1.AppearanceCell.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyTotalOpeningStockCalc1.AppearanceCell.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.colPropertyTotalOpeningStockCalc1.AppearanceCell.Options.UseFont = True
        Me.colPropertyTotalOpeningStockCalc1.AppearanceCell.Options.UseForeColor = True
        Me.colPropertyTotalOpeningStockCalc1.AppearanceCell.Options.UseTextOptions = True
        Me.colPropertyTotalOpeningStockCalc1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.BackColor = System.Drawing.Color.White
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.BorderColor = System.Drawing.Color.White
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.Options.UseBackColor = True
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.Options.UseBorderColor = True
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.Options.UseFont = True
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.Options.UseTextOptions = True
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
        Me.colPropertyTotalOpeningStockCalc1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.colPropertyTotalOpeningStockCalc1.DisplayFormat.FormatString = "N"
        Me.colPropertyTotalOpeningStockCalc1.FieldName = "PropertyTotalOpeningStockCalc"
        Me.colPropertyTotalOpeningStockCalc1.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Right
        Me.colPropertyTotalOpeningStockCalc1.MinWidth = 42
        Me.colPropertyTotalOpeningStockCalc1.Name = "colPropertyTotalOpeningStockCalc1"
        Me.colPropertyTotalOpeningStockCalc1.OptionsColumn.AllowEdit = False
        Me.colPropertyTotalOpeningStockCalc1.OptionsColumn.ReadOnly = True
        Me.colPropertyTotalOpeningStockCalc1.Visible = True
        Me.colPropertyTotalOpeningStockCalc1.VisibleIndex = 12
        Me.colPropertyTotalOpeningStockCalc1.Width = 143
        '
        'RepositoryItemComboBoxSOCIRent
        '
        Me.RepositoryItemComboBoxSOCIRent.AutoHeight = False
        Me.RepositoryItemComboBoxSOCIRent.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.RepositoryItemComboBoxSOCIRent.Items.AddRange(New Object() {"Social Rent", "Aff Rent"})
        Me.RepositoryItemComboBoxSOCIRent.Name = "RepositoryItemComboBoxSOCIRent"
        '
        'RepositoryItemComboBoxSOCIStockType
        '
        Me.RepositoryItemComboBoxSOCIStockType.Appearance.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepositoryItemComboBoxSOCIStockType.Appearance.Options.UseFont = True
        Me.RepositoryItemComboBoxSOCIStockType.AutoHeight = False
        Me.RepositoryItemComboBoxSOCIStockType.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.RepositoryItemComboBoxSOCIStockType.Items.AddRange(New Object() {"Gen Needs", "LCHO", "Supported", "Other", "Outright Sale", "Non-social"})
        Me.RepositoryItemComboBoxSOCIStockType.Name = "RepositoryItemComboBoxSOCIStockType"
        '
        'RepositoryItemTextEditThousandSep
        '
        Me.RepositoryItemTextEditThousandSep.AutoHeight = False
        Me.RepositoryItemTextEditThousandSep.DisplayFormat.FormatString = "#,###,##0"
        Me.RepositoryItemTextEditThousandSep.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.RepositoryItemTextEditThousandSep.Name = "RepositoryItemTextEditThousandSep"
        '
        'CardView1
        '
        Me.CardView1.GridControl = Me.GridControlStockGrid
        Me.CardView1.Name = "CardView1"
        '
        'LayoutViewStock
        '
        Me.LayoutViewStock.Columns.AddRange(New DevExpress.XtraGrid.Columns.LayoutViewColumn() {Me.LayoutViewColumn1})
        Me.LayoutViewStock.GridControl = Me.GridControlStockGrid
        Me.LayoutViewStock.Name = "LayoutViewStock"
        Me.LayoutViewStock.OptionsView.DefaultColumnCount = 2
        Me.LayoutViewStock.TemplateCard = Me.LayoutViewCard1
        '
        'LayoutViewColumn1
        '
        Me.LayoutViewColumn1.Caption = "LayoutViewColumn1"
        Me.LayoutViewColumn1.FieldName = "PropertyCurrentStockNumbers"
        Me.LayoutViewColumn1.LayoutViewField = Me.layoutViewField_LayoutViewColumn1
        Me.LayoutViewColumn1.MinWidth = 35
        Me.LayoutViewColumn1.Name = "LayoutViewColumn1"
        Me.LayoutViewColumn1.Width = 131
        '
        'layoutViewField_LayoutViewColumn1
        '
        Me.layoutViewField_LayoutViewColumn1.EditorPreferredWidth = 10
        Me.layoutViewField_LayoutViewColumn1.Location = New System.Drawing.Point(0, 0)
        Me.layoutViewField_LayoutViewColumn1.Name = "layoutViewField_LayoutViewColumn1"
        Me.layoutViewField_LayoutViewColumn1.Size = New System.Drawing.Size(203, 32)
        Me.layoutViewField_LayoutViewColumn1.TextSize = New System.Drawing.Size(177, 25)
        '
        'LayoutViewCard1
        '
        Me.LayoutViewCard1.HeaderButtonsLocation = DevExpress.Utils.GroupElementLocation.AfterText
        Me.LayoutViewCard1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.layoutViewField_LayoutViewColumn1})
        Me.LayoutViewCard1.Name = "layoutViewTemplateCard"
        '
        'TablePanelStockAssumptions
        '
        Me.TablePanelStockAssumptions.Appearance.BackColor = System.Drawing.Color.White
        Me.TablePanelStockAssumptions.Appearance.BorderColor = System.Drawing.Color.White
        Me.TablePanelStockAssumptions.Appearance.Options.UseBackColor = True
        Me.TablePanelStockAssumptions.Appearance.Options.UseBorderColor = True
        Me.TablePanelStockAssumptions.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 50.0!)})
        Me.TablePanelStockAssumptions.Controls.Add(Me.WindowsUIButtonPanelItemEdit)
        Me.TablePanelStockAssumptions.Controls.Add(Me.GridControlStockGrid)
        Me.TablePanelStockAssumptions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelStockAssumptions.Location = New System.Drawing.Point(0, 0)
        Me.TablePanelStockAssumptions.Margin = New System.Windows.Forms.Padding(8, 11, 8, 11)
        Me.TablePanelStockAssumptions.Name = "TablePanelStockAssumptions"
        Me.TablePanelStockAssumptions.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 85.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!)})
        Me.TablePanelStockAssumptions.Size = New System.Drawing.Size(1569, 882)
        Me.TablePanelStockAssumptions.TabIndex = 0
        Me.TablePanelStockAssumptions.UseSkinIndents = True
        '
        'WindowsUIButtonPanelItemEdit
        '
        Me.WindowsUIButtonPanelItemEdit.AppearanceButton.Normal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.WindowsUIButtonPanelItemEdit.AppearanceButton.Normal.Options.UseForeColor = True
        WindowsUIButtonImageOptions1.Image = CType(resources.GetObject("WindowsUIButtonImageOptions1.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions2.Image = CType(resources.GetObject("WindowsUIButtonImageOptions2.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions3.Image = CType(resources.GetObject("WindowsUIButtonImageOptions3.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions4.Image = CType(resources.GetObject("WindowsUIButtonImageOptions4.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions5.Image = CType(resources.GetObject("WindowsUIButtonImageOptions5.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions6.Image = CType(resources.GetObject("WindowsUIButtonImageOptions6.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions7.Image = CType(resources.GetObject("WindowsUIButtonImageOptions7.Image"), System.Drawing.Image)
        Me.WindowsUIButtonPanelItemEdit.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions1, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Close these assumptions", -1, True, Nothing, True, False, True, "Close", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(Nothing, True, -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions2, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Revert the assumptions to those from the last save", -1, True, Nothing, True, False, True, "Revert", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(Nothing, True, -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions3, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Apply assumptions to BP", -1, True, Nothing, True, False, True, "ApplyToFile", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions4, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Apply assumptions and save file", -1, True, Nothing, True, False, True, "ApplyAndSave", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions5, DevExpress.XtraBars.Docking2010.ButtonStyle.CheckButton, "Show extra additions and disposals", -1, True, Nothing, True, False, True, "AddDispo", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions6, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Add a new entry", -1, True, Nothing, True, False, True, "AddNew", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions7, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Delete this entry", -1, True, Nothing, True, False, True, "Delete", -1, False)})
        Me.TablePanelStockAssumptions.SetColumn(Me.WindowsUIButtonPanelItemEdit, 0)
        Me.WindowsUIButtonPanelItemEdit.ContentAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.WindowsUIButtonPanelItemEdit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowsUIButtonPanelItemEdit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.WindowsUIButtonPanelItemEdit.Location = New System.Drawing.Point(23, 20)
        Me.WindowsUIButtonPanelItemEdit.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.WindowsUIButtonPanelItemEdit.Name = "WindowsUIButtonPanelItemEdit"
        Me.WindowsUIButtonPanelItemEdit.Padding = New System.Windows.Forms.Padding(20, 0, 0, 0)
        Me.TablePanelStockAssumptions.SetRow(Me.WindowsUIButtonPanelItemEdit, 0)
        Me.WindowsUIButtonPanelItemEdit.Size = New System.Drawing.Size(1523, 79)
        Me.WindowsUIButtonPanelItemEdit.TabIndex = 2
        Me.WindowsUIButtonPanelItemEdit.Text = "WindowsUIButtonPanel1"
        '
        'StockAssumptionsInterface
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.TablePanelStockAssumptions)
        Me.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.Margin = New System.Windows.Forms.Padding(8, 11, 8, 11)
        Me.Name = "StockAssumptionsInterface"
        Me.Size = New System.Drawing.Size(1569, 882)
        CType(Me.UnboundSourceStocks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridControlStockGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridViewStockNumbers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemComboBoxOwnedManaged, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemLookUpEditSOCIStockType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemLookUpEditSOCIRentType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemIntegerEdit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEditInitialRateNewLettings, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemComboBoxSOCIRent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemComboBoxSOCIStockType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEditThousandSep, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CardView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutViewStock, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.layoutViewField_LayoutViewColumn1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutViewCard1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TablePanelStockAssumptions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelStockAssumptions.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UnboundSourceStocks As DevExpress.Data.UnboundSource
    Friend WithEvents GridControlStockGrid As DevExpress.XtraGrid.GridControl
    Friend WithEvents TablePanelStockAssumptions As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents GridViewStockNumbers As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents colPropertyStockDescription1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colPropertyOwnedManaged1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemComboBoxOwnedManaged As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents colPropertySOCIStockType1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemComboBoxSOCIStockType As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents colPropertySOCIRentType1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemComboBoxSOCIRent As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents colPropertyCurrentStockNumbers1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemIntegerEdit As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colPropertyPreBPlanStartDateNewBuild1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colPropertyPreBPlanStartDateDemolitions1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colPropertyPreBPlanStartDateRTBs1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colPropertyPreBPlanStartDateOtherDisposals1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colPropertyExistingStocksCalc1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colPropertyNewLettings1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colPropertyInitialRateNewLettings1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemSpinEditInitialRateNewLettings As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents colPropertyTotalOpeningStockCalc1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemLookUpEditSOCIRentType As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents RepositoryItemLookUpEditSOCIStockType As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents CardView1 As DevExpress.XtraGrid.Views.Card.CardView
    Friend WithEvents LayoutViewStock As DevExpress.XtraGrid.Views.Layout.LayoutView
    Friend WithEvents LayoutViewCard1 As DevExpress.XtraGrid.Views.Layout.LayoutViewCard
    Friend WithEvents LayoutViewColumn1 As DevExpress.XtraGrid.Columns.LayoutViewColumn
    Friend WithEvents layoutViewField_LayoutViewColumn1 As DevExpress.XtraGrid.Views.Layout.LayoutViewField
    Friend WithEvents WindowsUIButtonPanelItemEdit As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
    Friend WithEvents RepositoryItemTextEditThousandSep As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
End Class
