param(
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
$structurePath = Join-Path $PSScriptRoot '..\Structure.xml'

function XmlText([string]$Value) {
    if ($null -eq $Value) { return '' }
    return [Security.SecurityElement]::Escape($Value)
}

function New-OutputSection([pscustomobject]$Section) {
    $type = $Section.Type
    $dataRange = $Section.DataRange
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add("      <CSInterfaceSection Name=`"$(XmlText $Section.Name)`">")
    $lines.Add("        <ISName>$(XmlText $Section.Name)</ISName>")
    $lines.Add('        <ISDatasource>')
    $lines.Add("          <ISDName>$(XmlText $Section.Name)</ISDName>")
    $lines.Add('          <ISDID>0</ISDID>')
    $lines.Add("          <DSType>$type</DSType>")
    $lines.Add('          <DSSource>CR</DSSource>')
    $lines.Add('          <MergeSources>None</MergeSources>')
    $lines.Add('          <RO>TRUE</RO>')
    $lines.Add('          <MergedHeader>False</MergedHeader>')
    $lines.Add('          <SourceDataFormat>Cols</SourceDataFormat>')
    $lines.Add('          <Pivot>TRUE</Pivot>')
    $lines.Add('          <ColsExpandBy>None</ColsExpandBy>')
    $lines.Add('          <RowsExpandBy>None</RowsExpandBy>')
    $lines.Add('          <CellRangeDataSource>')
    $lines.Add("            <Worksheet>$(XmlText $Section.Worksheet)</Worksheet>")
    $lines.Add('            <PositID>1</PositID>')
    $lines.Add("            <DSType>$type</DSType>")
    $lines.Add('            <ColsDefinedBy>None</ColsDefinedBy>')
    $lines.Add('            <RowsDefinedBy>None</RowsDefinedBy>')
    $lines.Add("            <DataRange>$(XmlText $dataRange)</DataRange>")
    if ($Section.SourceName) { $lines.Add("            <LiveGridSourceName>$(XmlText $Section.SourceName)</LiveGridSourceName>") }
    if ($Section.SourceRanges) { $lines.Add("            <LiveGridSourceRanges>$(XmlText $Section.SourceRanges)</LiveGridSourceRanges>") }
    if ($Section.Areas) { $lines.Add("            <LiveGridSourceAreas>$(XmlText $Section.Areas)</LiveGridSourceAreas>") }
    if ($Section.Leading) { $lines.Add("            <LiveGridLeadingColumns>$(XmlText $Section.Leading)</LiveGridLeadingColumns>") }
    if ($Section.Headers) { $lines.Add("            <LiveGridHeaderRows>$(XmlText $Section.Headers)</LiveGridHeaderRows>") }
    if ($Section.CategoryRow) { $lines.Add("            <LiveVGridCategoryRow>$(XmlText $Section.CategoryRow)</LiveVGridCategoryRow>") }
    if ($Section.RecordHeaders) { $lines.Add("            <LiveVGridRecordHeaderColumns>$(XmlText $Section.RecordHeaders)</LiveVGridRecordHeaderColumns>") }
    $lines.Add('            <Pivot>TRUE</Pivot>')
    $lines.Add('            <DataFieldDefinition>')
    $lines.Add('              <FieldName>Workbook value</FieldName>')
    $lines.Add('              <DataFormat>S</DataFormat>')
    $lines.Add('              <Summary>None</Summary>')
    $lines.Add('            </DataFieldDefinition>')
    $lines.Add('          </CellRangeDataSource>')
    $lines.Add('        </ISDatasource>')
    $lines.Add('        <IElement>')
    $lines.Add("          <Type>$type</Type>")
    $lines.Add('          <DataSource>0</DataSource>')
    $lines.Add('        </IElement>')
    $lines.Add('      </CSInterfaceSection>')
    return $lines
}

function S($Name, $Worksheet, $Type, $DataRange, $SourceName = '', $Areas = '', $Leading = '', $Headers = '', $CategoryRow = '', $RecordHeaders = '2', $SourceRanges = '') {
    [pscustomobject]@{ Name=$Name; Worksheet=$Worksheet; Type=$Type; DataRange=$DataRange; SourceName=$SourceName; Areas=$Areas; Leading=$Leading; Headers=$Headers; CategoryRow=$CategoryRow; RecordHeaders=$RecordHeaders; SourceRanges=$SourceRanges }
}

$children = @(
    [pscustomobject]@{ Id=1; Name='Development Stock Numbers'; Group='Stock Numbers'; Sheet='Development Stock Numbers'; Sections=@(
        (S 'Development Opening Stock' 'Development Stock Numbers' 'LiveGrid' 'A11:B51' 'Outputs_DevelopmentStockNumbers' '2' '2' '6,9' '' ''),
        (S 'New Development Units' 'Development Stock Numbers' 'LiveGrid' 'A11:B51' 'Outputs_DevelopmentStockNumbers' '3,4' '2' '6,9' '' ''),
        (S 'Development Staircasing' 'Development Stock Numbers' 'LiveGrid' 'A11:B51' 'Outputs_DevelopmentStockNumbers' '5' '2' '6,9' '' '')
    )},
    [pscustomobject]@{ Id=2; Name='Stock Numbers'; Group='Stock Numbers'; Sheet='Stock Numbers'; Sections=@(
        (S 'Consolidated Opening Stock Numbers' 'Stock Numbers' 'LiveGrid' 'A12:AH52' 'Outputs_StockNumbers' '1' '' '6,10' '' ''),
        (S 'All Stock Numbers' 'Stock Numbers' 'LiveGrid' 'A12:B52' 'Outputs_StockNumbers' '2' '2' '6,10' '' ''),
        (S 'Managed Stock Numbers' 'Stock Numbers' 'LiveGrid' 'A12:B52' 'Outputs_StockNumbers' '3' '2' '6,10' '' ''),
        (S 'Owned Stock Numbers' 'Stock Numbers' 'LiveGrid' 'A12:B52' 'Outputs_StockNumbers' '4' '2' '6,10' '' '')
    )},
    [pscustomobject]@{ Id=3; Name='Existing Cashflows'; Group='Cashflows'; Sheet='Existing Cashflows'; Sections=@((S 'Existing Cashflows' 'Existing Cashflows' 'LiveGrid' 'A9:AA48' 'Outputs_ExistingCashflows' '' '' '6,7' '' '')) },
    [pscustomobject]@{ Id=4; Name='Development Cashflows'; Group='Cashflows'; Sheet='Development Cashflows'; Sections=@((S 'Development Cashflows' 'Development Cashflows' 'LiveGrid' 'A10:H49' 'Outputs_DevelopmentCashflows' '' '' '6,8' '' '')) },
    [pscustomobject]@{ Id=5; Name='Cashflow'; Group='Cashflows'; Sheet='Cashflow'; Sections=@((S 'Cashflow' 'Cashflow' 'LiveGrid' 'A11:AA50' 'Outputs_Cashflow' '' '' '6,9' '' '')) },
    [pscustomobject]@{ Id=6; Name='Existing Cashflows Detailed'; Group='Cashflows'; Sheet='Existing Cashflows Detailed'; Sections=@((S 'Existing Cashflows Detailed' 'Existing Cashflows Detailed' 'LiveGrid' 'A9:AD48' 'Outputs_ExistingCashflowsDetailed' '' '' '5,7' '' '')) },
    [pscustomobject]@{ Id=7; Name='Development Cashflow Detailed'; Group='Cashflows'; Sheet='Development Cashflow Detailed'; Sections=@((S 'Development Cashflow Detailed' 'Development Cashflow Detailed' 'LiveGrid' 'A9:AY48' 'Outputs_DevelopmentCashflowDetailed' '' '' '5,7' '' '')) },
    [pscustomobject]@{ Id=8; Name='Cashflow Detail before Interco'; Group='Cashflows'; Sheet='Cashflow det b4 interco'; Sections=@((S 'Cashflow Detail before Interco' 'Cashflow det b4 interco' 'LiveGrid' 'A9:AD48' 'Outputs_Cashflowdetb4interco' '' '' '5,7' '' '')) },
    [pscustomobject]@{ Id=9; Name='Cashflow Interco'; Group='Cashflows'; Sheet='Cashflow interco'; Sections=@((S 'Cashflow Interco' 'Cashflow interco' 'LiveGrid' 'A9:AD48' '' '' '' '5,7' '' '' 'A9:AD48;AF9:AR48;AT9:AV48')) },
    [pscustomobject]@{ Id=10; Name='Cash Journals Detailed'; Group='Cashflows'; Sheet='Cash Journals Detailed'; Sections=@((S 'Cash Journals Detailed' 'Cash Journals Detailed' 'LiveGrid' 'A9:AG48' 'Outputs_CashJournalsDetailed' '' '' '5,7' '' '')) },
    [pscustomobject]@{ Id=11; Name='Cashflow Detailed'; Group='Cashflows'; Sheet='Cashflow detailed'; Sections=@((S 'Cashflow Detailed' 'Cashflow detailed' 'LiveGrid' 'A9:AD48' 'Outputs_Cashflowdetailed' '' '' '5,7' '' '')) },
    [pscustomobject]@{ Id=12; Name='Summary Comp Inc - Trad View'; Group='Accounts'; Sheet='Summary Comp Inc - Trad View'; Sections=@((S 'Summary Comp Inc - Trad View' 'Summary Comp Inc - Trad View' 'LiveGrid' 'A9:AP43' '' '' '' '5,6,7' '' '' 'A9:AP43')) },
    [pscustomobject]@{ Id=13; Name='Detailed Comp Inc - Trad View'; Group='Accounts'; Sheet='Detailed Comp Inc - Trad View'; Sections=@((S 'Detailed Comp Inc - Trad View' 'Detailed Comp Inc - Trad View' 'LiveGrid' 'A9:AP97' '' '' '' '5,6,7' '' '' 'A9:AP97')) },
    [pscustomobject]@{ Id=14; Name='Financial Position - Trad View'; Group='Accounts'; Sheet='Financial Position - Trad View'; Sections=@((S 'Financial Position - Trad View' 'Financial Position - Trad View' 'LiveGrid' 'A9:AQ67' '' '' '' '5,6,7' '' '' 'A9:AQ67')) },
    [pscustomobject]@{ Id=15; Name='Cashflow Statement - Trad View'; Group='Accounts'; Sheet='Cashflow Statement - Trad View'; Sections=@((S 'Cashflow Statement - Trad View' 'Cashflow Statement - Trad View' 'LiveGrid' 'A9:AP67' '' '' '' '5,6,7' '' '' 'A9:AP67')) },
    [pscustomobject]@{ Id=16; Name='Summary Comp Inc - Alt View'; Group='Accounts'; Sheet='Summary Comp Inc - Alt View'; Sections=@((S 'Summary Comp Inc - Alt View' 'Summary Comp Inc - Alt View' 'LiveGrid' 'A9:AK48' '' '' '' '6,7' '' '' 'A9:AK48')) },
    [pscustomobject]@{ Id=17; Name='Detailed Comp Inc - Alt View'; Group='Accounts'; Sheet='Detailed Comp Inc - Alt View'; Sections=@((S 'Detailed Comp Inc - Alt View' 'Detailed Comp Inc - Alt View' 'LiveGrid' 'A9:CQ48' '' '' '' '6,7' '' '' 'A9:CQ48')) },
    [pscustomobject]@{ Id=18; Name='Financial Position - Alt View'; Group='Accounts'; Sheet='Financial Position - Alt View'; Sections=@((S 'Financial Position - Alt View' 'Financial Position - Alt View' 'LiveGrid' 'A10:BE50' '' '' '' '7,8' '' '' 'A10:BE50')) },
    [pscustomobject]@{ Id=19; Name='Cashflow Statement - Alt View'; Group='Accounts'; Sheet='Cashflow Statement - Alt View'; Sections=@((S 'Cashflow Statement - Alt View' 'Cashflow Statement - Alt View' 'LiveGrid' 'A9:BE48' '' '' '' '6,7' '' '' 'A9:BE48')) }
)

$additionalChildren = @(
    [pscustomobject]@{ Id=22; Name='Covenants'; Group='Covenants and Other Reports'; Sheet='Covenants'; Sections=@(
        (S 'Covenants' 'Covenants' 'LiveGrid' 'A11:E50' 'Outputs_Covenants' '' '' '6,7,8,9' '' '')
    )},
    [pscustomobject]@{ Id=23; Name='Value for Money Metrics'; Group='Covenants and Other Reports'; Sheet='Value for Money Metrics'; Sections=@(
        (S 'Value for Money Metrics' 'Value for Money Metrics' 'LiveGrid' 'A9:M171' '' '' '' '5,6' '' '' 'A9:M171')
    )},
    [pscustomobject]@{ Id=24; Name='Loan Output Table'; Group='Covenants and Other Reports'; Sheet='Loan Output Table'; Sections=@(
        (S 'Loan Output Table' 'Loan Output Table' 'LiveGrid' 'C10:K49' 'Outputs_LoanOutputTable' '' '' '5,6,7,8' '' '')
    )},
    [pscustomobject]@{ Id=25; Name='Summary Stock Numbers'; Group='Covenants and Other Reports'; Sheet='Summary Stock Numbers'; Sections=@(
        (S 'Summary Stock Numbers' 'Summary Stock Numbers' 'LiveGrid' 'A10:B50' 'Outputs_SummaryStockNumbers' '' '' '5,6,7,8' '' '')
    )},
    [pscustomobject]@{ Id=26; Name='Surplus on Sales'; Group='Covenants and Other Reports'; Sheet='Surplus on Sales'; Sections=@(
        (S 'Surplus on Sales' 'Surplus on Sales' 'LiveGrid' 'A9:B48' 'Outputs_SurplusonSales' '' '' '5,6,7' '' '')
    )},
    [pscustomobject]@{ Id=27; Name='NSH Breakdown'; Group='Covenants and Other Reports'; Sheet='NSH Breakdown'; Sections=@(
        (S 'NSH Breakdown' 'NSH Breakdown' 'LiveGrid' 'A10:B49' 'Outputs_NSHBreakdown' '' '' '5,6,7,8' '' '')
    )},
    [pscustomobject]@{ Id=28; Name='Hsg Properties Mvmt'; Group='Covenants and Other Reports'; Sheet='Hsg Properties Mvmt'; Sections=@(
        (S 'Hsg Properties Mvmt' 'Hsg Properties Mvmt' 'LiveGrid' 'A10:B49' 'Outputs_HsgPropertiesMvmt' '' '' '5,6,7' '' '')
    )},
    [pscustomobject]@{ Id=29; Name='BP Input Scheme Cashflows'; Group='Other Cashflow Outputs'; Sheet='BP Input Scheme Cashflows'; Sections=@(
        (S 'BP Input Scheme Cashflows' 'BP Input Scheme Cashflows' 'LiveGrid' 'A12:Z106' 'Outputs_BPInputSchemeCashflows' '' '' '7,8,9,10' '' '')
    )},
    [pscustomobject]@{ Id=30; Name='5 Yr Monthly Cashflow'; Group='Other Cashflow Outputs'; Sheet='5 Yr Monthly Cashflow'; Sections=@(
        (S '5 Yr Monthly Cashflow' '5 Yr Monthly Cashflow' 'LiveGrid' 'A8:F102' 'Outputs_5YrMonthlyCashflow' '' '' '5,6' '' '')
    )},
    [pscustomobject]@{ Id=31; Name='5 Yr Quarterly Cashflow'; Group='Other Cashflow Outputs'; Sheet='5 Yr Quarterly Cashflow'; Sections=@(
        (S '5 Yr Quarterly Cashflow' '5 Yr Quarterly Cashflow' 'LiveGrid' 'A8:E27' 'Outputs_5YrQuarterlyCashflow' '' '' '5,6' '' '')
    )}
)

$text = [IO.File]::ReadAllText($structurePath)
$existing = [regex]::Match($text, '(?s)    <ChildStructure Name="Existing Stock Numbers">.*?    </ChildStructure>').Value
if (-not $existing) { throw 'Existing Stock Numbers Output child was not found.' }

$group = [Collections.Generic.List[string]]::new()
$group.Add('  <GroupStructure Name="Outputs">')
$group.Add('    <GSName>Outputs</GSName>')
$group.Add('    <GSID>2</GSID>')
$group.Add('    <FirstChild>0</FirstChild>')
$group.Add($existing)

foreach ($child in $children) {
    $group.Add("    <ChildStructure Name=`"$(XmlText $child.Name)`">")
    $group.Add("      <CSName>$(XmlText $child.Name)</CSName>")
    $group.Add("      <CSID>$($child.Id)</CSID>")
    $group.Add('      <ParentID>2</ParentID>')
    $group.Add('      <IsMaster>False</IsMaster>')
    $group.Add("      <GroupName>$(XmlText $child.Group)</GroupName>")
    $group.Add("      <DefaultWorksheet>$(XmlText $child.Sheet)</DefaultWorksheet>")
    foreach ($section in $child.Sections) {
        foreach ($line in (New-OutputSection $section)) { $group.Add($line) }
    }
    $group.Add('    </ChildStructure>')
}

$group.Add('    <ChildStructure Name="BP Dashboard">')
$group.Add('      <CSName>BP Dashboard</CSName>')
$group.Add('      <CSID>20</CSID>')
$group.Add('      <ParentID>2</ParentID>')
$group.Add('      <IsMaster>False</IsMaster>')
$group.Add('      <SpecialElement>BP_Dashboard</SpecialElement>')
$group.Add('      <GroupName>Dashboards</GroupName>')
$group.Add('      <DefaultWorksheet>BP Dashboard</DefaultWorksheet>')
$group.Add('      <CSInterfaceSection Name="BPDashboard">')
$group.Add('        <ISName>BPDashboard</ISName>')
$group.Add('        <IElement>')
$group.Add('          <Type>Interface</Type>')
$group.Add('          <DataSource>BP_Dashboard</DataSource>')
$group.Add('        </IElement>')
$group.Add('      </CSInterfaceSection>')
$group.Add('    </ChildStructure>')
$group.Add('    <ChildStructure Name="Funding Dashboard">')
$group.Add('      <CSName>Funding Dashboard</CSName>')
$group.Add('      <CSID>21</CSID>')
$group.Add('      <ParentID>2</ParentID>')
$group.Add('      <IsMaster>False</IsMaster>')
$group.Add('      <SpecialElement>FundingDashboard</SpecialElement>')
$group.Add('      <GroupName>Dashboards</GroupName>')
$group.Add('      <DefaultWorksheet>Funding Dashboard</DefaultWorksheet>')
$group.Add('      <CSInterfaceSection Name="Funding Dashboard">')
$group.Add('        <ISName>Funding Dashboard</ISName>')
$group.Add('        <IElement>')
$group.Add('          <Type>Interface</Type>')
$group.Add('          <DataSource>Funding_Dashboard</DataSource>')
$group.Add('        </IElement>')
$group.Add('      </CSInterfaceSection>')
$group.Add('    </ChildStructure>')
foreach ($child in $additionalChildren) {
    $group.Add(('    <ChildStructure Name="{0}">' -f (XmlText $child.Name)))
    $group.Add(('      <CSName>{0}</CSName>' -f (XmlText $child.Name)))
    $group.Add(('      <CSID>{0}</CSID>' -f $child.Id))
    $group.Add('      <ParentID>2</ParentID>')
    $group.Add('      <IsMaster>False</IsMaster>')
    $group.Add(('      <GroupName>{0}</GroupName>' -f (XmlText $child.Group)))
    $group.Add(('      <DefaultWorksheet>{0}</DefaultWorksheet>' -f (XmlText $child.Sheet)))
    foreach ($section in $child.Sections) {
        foreach ($line in (New-OutputSection $section)) { $group.Add($line) }
    }
    $group.Add('    </ChildStructure>')
}
$group.Add('    <ChildStructure Name="Analysis V1">')
$group.Add('      <CSName>Analysis V1</CSName>')
$group.Add('      <CSID>32</CSID>')
$group.Add('      <ParentID>2</ParentID>')
$group.Add('      <IsMaster>False</IsMaster>')
$group.Add('      <SpecialElement>BPIncomeExpenditureAnalyser</SpecialElement>')
$group.Add('      <GroupName>Analysis</GroupName>')
$group.Add('      <CSInterfaceSection Name="Analysis V1">')
$group.Add('        <ISName>Analysis V1</ISName>')
$group.Add('        <IElement>')
$group.Add('          <Type>Interface</Type>')
$group.Add('          <DataSource>BPIncomeExpenditureAnalyser</DataSource>')
$group.Add('        </IElement>')
$group.Add('      </CSInterfaceSection>')
$group.Add('    </ChildStructure>')
$group.Add('    <ChildStructure Name="Analysis V2">')
$group.Add('      <CSName>Analysis V2</CSName>')
$group.Add('      <CSID>33</CSID>')
$group.Add('      <ParentID>2</ParentID>')
$group.Add('      <IsMaster>False</IsMaster>')
$group.Add('      <SpecialElement>BPIncomeExpenditureAnalyserV2</SpecialElement>')
$group.Add('      <GroupName>Analysis</GroupName>')
$group.Add('      <CSInterfaceSection Name="Analysis V2">')
$group.Add('        <ISName>Analysis V2</ISName>')
$group.Add('        <IElement>')
$group.Add('          <Type>Interface</Type>')
$group.Add('          <DataSource>BPIncomeExpenditureAnalyserV2</DataSource>')
$group.Add('        </IElement>')
$group.Add('      </CSInterfaceSection>')
$group.Add('    </ChildStructure>')
$group.Add('  </GroupStructure>')

$replacement = $group -join "`r`n"
$updated = [regex]::Replace($text, '(?s)  <GroupStructure Name="Outputs">.*?  </GroupStructure>', [Text.RegularExpressions.MatchEvaluator]{ param($m) $replacement }, 1)
if ($Apply) {
    if ($updated -eq $text) {
        Write-Output 'Structure.xml Outputs children 0-33 (excluding Scenario Planning) are already current.'
        return
    }
    [IO.File]::WriteAllText($structurePath, $updated, [Text.UTF8Encoding]::new($false))
    Write-Output 'Updated Structure.xml Outputs children 0-33 (excluding Scenario Planning).'
} else {
    Write-Output $replacement
}
