param([switch]$Apply)

$ErrorActionPreference = 'Stop'

function Add-WorkingsItems {
    param([System.Collections.Generic.List[object]]$Target, [string]$Group, [string[]]$Items)
    foreach ($Item in $Items) {
        $Parts = $Item.Split('|', 2)
        $Target.Add([pscustomobject]@{
            Group = $Group
            Name = $Parts[0]
            Worksheet = if ($Parts.Count -gt 1) { $Parts[1] } else { $Parts[0] }
        })
    }
}

function ConvertTo-XmlText {
    param([string]$Value)
    return [System.Security.SecurityElement]::Escape($Value)
}

function Convert-LetterSuffixToNumber {
    param([string]$Suffix)
    if ($Suffix -notmatch '^[A-Z]+$') { return [int]::MaxValue }
    $Result = 0
    foreach ($Character in $Suffix.ToCharArray()) {
        $Result = ($Result * 26) + ([int]$Character - [int][char]'A' + 1)
    }
    return $Result
}

function Get-ShortDefinedName {
    param([string]$Name)
    return ($Name -replace '^.*!', '')
}

function Get-NameWorksheet {
    param([string]$RefersTo)
    $Match = [regex]::Match($RefersTo, "^='((?:[^']|'')+)'!")
    if ($Match.Success) { return $Match.Groups[1].Value.Replace("''", "'") }
    $Match = [regex]::Match($RefersTo, '^=([^!]+)!')
    if ($Match.Success) { return $Match.Groups[1].Value }
    return $null
}

function Get-NameRangeInfo {
    param($Workbook, $DefinedName, [string]$WorksheetName)
    $Worksheet = $Workbook.Worksheets.Item($WorksheetName)
    $Ranges = New-Object System.Collections.Generic.List[object]
    $RefersTo = ([string]$DefinedName.RefersTo).TrimStart('=')

    foreach ($Part in $RefersTo.Split(',')) {
        $BangIndex = $Part.LastIndexOf('!')
        if ($BangIndex -lt 0) { continue }
        $Address = $Part.Substring($BangIndex + 1).Replace('$', '').Trim()
        try {
            $Range = $Worksheet.Range($Address)
            $Ranges.Add([pscustomobject]@{
                Address = $Range.Address($false, $false, 1, $false)
                Top = [int]$Range.Row
                Bottom = [int]$Range.Row + [int]$Range.Rows.Count - 1
                Left = [int]$Range.Column
                Right = [int]$Range.Column + [int]$Range.Columns.Count - 1
                RowCount = [int]$Range.Rows.Count
            })
        }
        catch {
            continue
        }
    }

    if ($Ranges.Count -eq 0) { return $null }
    $SelectedTop = ($Ranges | Measure-Object Top -Maximum).Maximum
    $AtSelectedTop = @($Ranges | Where-Object Top -eq $SelectedTop)
    $SelectedRowCount = $AtSelectedTop |
        Group-Object RowCount |
        Sort-Object Count, Name -Descending |
        Select-Object -First 1 |
        ForEach-Object { [int]$_.Name }
    $Selected = @($AtSelectedTop | Where-Object RowCount -eq $SelectedRowCount)

    return [pscustomobject]@{
        Ranges = $Selected
        Top = [int]$SelectedTop
        FirstAddress = $Selected[0].Address
        Left = ($Selected | Measure-Object Left -Minimum).Minimum
        Right = ($Selected | Measure-Object Right -Maximum).Maximum
    }
}

function Get-BlockTitle {
    param($Workbook, [string]$WorksheetName, $RangeInfo, [string]$Fallback)
    if ($null -eq $RangeInfo) { return $Fallback }
    $Worksheet = $Workbook.Worksheets.Item($WorksheetName)
    $MaximumColumn = [Math]::Min(8, [Math]::Max(1, [int]$RangeInfo.Left + 3))
    $MinimumRow = [Math]::Max(1, [int]$RangeInfo.Top - 8)
    $SearchColumns = New-Object System.Collections.Generic.List[int]

    if ($RangeInfo.Ranges.Count -gt 1 -and
        ($RangeInfo.Ranges[0].Right - $RangeInfo.Ranges[0].Left + 1) -le 3) {
        $FirstResultArea = $RangeInfo.Ranges[1]
        for ($Column = $FirstResultArea.Left;
             $Column -le [Math]::Min($FirstResultArea.Right, $FirstResultArea.Left + 3);
             $Column++) {
            $SearchColumns.Add($Column)
        }
    }
    for ($Column = 1; $Column -le $MaximumColumn; $Column++) {
        if (-not $SearchColumns.Contains($Column)) { $SearchColumns.Add($Column) }
    }

    for ($Row = [int]$RangeInfo.Top - 1; $Row -ge $MinimumRow; $Row--) {
        foreach ($Column in $SearchColumns) {
            $Text = ([string]$Worksheet.Cells.Item($Row, $Column).Text).Trim()
            if ([string]::IsNullOrWhiteSpace($Text)) { continue }
            if ($Text -match '^(Year|Date From|Date To|Period|Total|Days?|�|%|Go to Contents)$') { continue }
            if ($Text -match '^[#0-9.,%() /:-]+$') { continue }
            if ($Text.Length -lt 3) { continue }
            return $Text
        }
    }
    return $Fallback
}

function Get-DirectRangeInfo {
    param($Workbook, [string]$WorksheetName)
    $Worksheet = $Workbook.Worksheets.Item($WorksheetName)
    $Used = $Worksheet.UsedRange
    $Top = [int]$Used.Row
    $Left = [int]$Used.Column
    $Bottom = $Top + [int]$Used.Rows.Count - 1
    $Right = $Left + [int]$Used.Columns.Count - 1
    $SearchBottom = [Math]::Min($Bottom, $Top + 20)
    $BestRow = $Top
    $BestScore = -1

    for ($Row = $Top; $Row -le $SearchBottom; $Row++) {
        $NonEmpty = 0
        $TextValues = 0
        for ($Column = $Left; $Column -le $Right; $Column++) {
            $Value = $Worksheet.Cells.Item($Row, $Column).Value2
            if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { continue }
            $NonEmpty++
            if ($Value -is [string]) { $TextValues++ }
        }
        $Score = ($TextValues * 3) + $NonEmpty
        if ($Score -gt $BestScore) {
            $BestScore = $Score
            $BestRow = $Row
        }
    }

    $DataTop = if ($BestRow -lt $Bottom) { $BestRow + 1 } else { $Top }
    $DataRange = $Worksheet.Range(
        $Worksheet.Cells.Item($DataTop, $Left),
        $Worksheet.Cells.Item($Bottom, $Right))

    return [pscustomobject]@{
        Address = $DataRange.Address($false, $false, 1, $false)
        HeaderRow = $BestRow
    }
}

function Get-AlternativeLabel {
    param([string]$Name, [string]$Fallback)
    if ($Name -eq 'StressLiveInfo') { return 'Live Stress Information' }
    if ($Name -match '^S0Data$') { return 'Base Case' }
    if ($Name -match '^S([0-9]+)Data$') { return 'Test ' + $Matches[1] }
    return $Fallback
}

function New-LiveGridSectionXml {
    param(
        [string]$SectionName,
        [string]$Worksheet,
        [string]$SourceName,
        [string]$SourceRanges,
        [string]$DataRange,
        [string]$HeaderRows
    )

    $EscapedSection = ConvertTo-XmlText $SectionName
    $EscapedWorksheet = ConvertTo-XmlText $Worksheet
    $EscapedDataRange = ConvertTo-XmlText $DataRange
    $SourceXml = if ([string]::IsNullOrWhiteSpace($SourceName)) {
        "            <LiveGridSourceRanges>$(ConvertTo-XmlText $SourceRanges)</LiveGridSourceRanges>"
    }
    else {
        "            <LiveGridSourceName>$(ConvertTo-XmlText $SourceName)</LiveGridSourceName>"
    }
    $HeaderXml = if ([string]::IsNullOrWhiteSpace($HeaderRows)) {
        ''
    }
    else {
        "            <LiveGridHeaderRows>$(ConvertTo-XmlText $HeaderRows)</LiveGridHeaderRows>" + [Environment]::NewLine
    }

    return @"
      <CSInterfaceSection Name="$EscapedSection">
        <ISName>$EscapedSection</ISName>
        <ISDatasource>
          <ISDName>$EscapedSection</ISDName>
          <ISDID>0</ISDID>
          <DSType>LiveGrid</DSType>
          <DSSource>CR</DSSource>
          <MergeSources>None</MergeSources>
          <RO>TRUE</RO>
          <MergedHeader>False</MergedHeader>
          <SourceDataFormat>Cols</SourceDataFormat>
          <Pivot>TRUE</Pivot>
          <ColsExpandBy>None</ColsExpandBy>
          <RowsExpandBy>None</RowsExpandBy>
          <CellRangeDataSource>
            <Worksheet>$EscapedWorksheet</Worksheet>
            <PositID>1</PositID>
            <DSType>LiveGrid</DSType>
            <ColsDefinedBy>None</ColsDefinedBy>
            <RowsDefinedBy>None</RowsDefinedBy>
            <DataRange>$EscapedDataRange</DataRange>
$SourceXml
$HeaderXml            <Pivot>TRUE</Pivot>
            <DataFieldDefinition>
              <FieldName>Value</FieldName>
              <DataFormat>S</DataFormat>
            </DataFieldDefinition>
          </CellRangeDataSource>
        </ISDatasource>
        <IElement>
          <Type>LiveGrid</Type>
          <DataSource>0</DataSource>
        </IElement>
      </CSInterfaceSection>
"@
}

function New-ChildXml {
    param([int]$Index, $Definition, [string[]]$Sections)
    $Name = ConvertTo-XmlText $Definition.Name
    $Group = ConvertTo-XmlText $Definition.Group
    $Worksheet = ConvertTo-XmlText $Definition.Worksheet
    $SectionText = $Sections -join ''

    return @"
    <ChildStructure Name="$Name">
      <CSName>$Name</CSName>
      <CSID>$Index</CSID>
      <ParentID>1</ParentID>
      <IsMaster>False</IsMaster>
      <GroupName>$Group</GroupName>
      <DefaultWorksheet>$Worksheet</DefaultWorksheet>
$SectionText    </ChildStructure>
"@
}

$Definitions = New-Object System.Collections.Generic.List[object]
Add-WorkingsItems $Definitions 'Stock' @(
    'Target Rent Letting Rates', 'Target Rent Letting Numbers',
    'Demolition Numbers', 'RTB Numbers', 'Other Disposal Numbers',
    'Stock Conversion Numbers', 'Development Stock',
    'Development Stock Numbers', 'Service Charge Numbers', 'Leaseholder Units'
)
Add-WorkingsItems $Definitions 'Rental & Service Charge Income' @(
    'Rent Factors', 'Service Charge Factors', 'Void Rates', 'Bad Debts Rates',
    'Existing Unit Rents', 'Unit Service Charges', 'Existing Rental Income',
    'Existing Rental Void Losses', 'Existing Rental Bad Debts',
    'Service Charge Income', 'Service Charge Voids', 'Service Charge Bad Debts',
    'Arrears'
)
Add-WorkingsItems $Definitions 'Disposals & Other Income' @(
    'RTB Valuation Factors', 'Other Income Factors', 'RTB Income Foregone',
    'RTB Receipts', 'Other Disposal Income', 'Owner Occupier Income',
    'Unit Leaseholder Charges', 'Leaseholder Income', 'Specific Income Drivers',
    'Specific Income Workings', 'Specific Income Voids',
    'Specific Income Bad Debts', 'Other Income Workings',
    'Capital Grant Workings', 'Intercompany Income'
)
Add-WorkingsItems $Definitions 'Management Costs' @(
    'Management Cost Drivers', 'Management Cost Factors',
    'Fixed Management Costs', 'Variable Management Costs',
    'Existing Management Costs'
)
Add-WorkingsItems $Definitions 'Repairs and Maintenance' @(
    'Stock Condition Results', 'Repairs & Maint. Drivers',
    'Repairs & Maint. Cost Factors', 'Repairs & Maintenance Costs',
    'Repairs & Maintenance Depn', 'Cost & Depn on Replacement',
    'Existing Repairs & Maint. Costs'
)
Add-WorkingsItems $Definitions 'Other Costs and Capital Items' @(
    'JV Opening balances', 'JV Investments', 'JV Share of Profits',
    'JV Payments Repayments', 'JV Closing Balances',
    'JV Interest Received', 'Interco Expenditure'
)
Add-WorkingsItems $Definitions 'Development' @(
    'Development Inflation Factors', 'Development Stock',
    'Development Capital', 'Development Revenue', 'Development Expenditure',
    'Dvpt NonCash', 'Dvpt Component Depn', 'Dvpt Maint Capitalisation'
)
Add-WorkingsItems $Definitions 'Development Import' @(
    'Dvpt Factors', 'Dvpt Imp Cashflows', 'Dvpt Imp Non Cash',
    'All Schemes BP Cashflow', 'Dvpt Imp CF Adj', 'All Schemes Cashflow',
    'Non Tenure Capital Expenditure', 'All Schemes BP Non Cash',
    'Dvpt Imp NC Adj', 'All Schemes Non Cash', 'All Schemes FFR',
    'Dvpt Imp FFR Info', 'Dvpt Cashflows'
)
Add-WorkingsItems $Definitions 'Funding & Covenants' @(
    'Funding Flowchart', 'Interest Rates', 'Loan Interest Payable',
    'Loan Opening Balances', 'Loan Drawdowns', 'Loan Interest Paid',
    'Loan Repayments', 'Loan Closing Balances', 'Loan Closing Facilities',
    'Loan Fixed Variable', 'Loan Commitment Fees', 'Loan Fees Amortisation',
    'Bond Premium Amortisation', 'Cash on Deposit', 'Investments',
    'InterCo Opening Balances', 'InterCo Increases', 'InterCo Decreases',
    'InterCo Closing Balances', 'InterCo Interest', 'Taxation Computation'
)
Add-WorkingsItems $Definitions 'Component Accounting' @(
    'Depn Type Stock Numbers', 'Component 1', 'Component 2', 'Component 3',
    'Component 4', 'Component 5', 'Component 6', 'Component 7',
    'Component 8', 'Component 9', 'Component 10', 'Component 11',
    'Component 12', 'Component Totals'
)
Add-WorkingsItems $Definitions 'Accounts' @(
    'Depreciation Capitalised Costs', 'OFA Additions', 'OFA Workings',
    'Opening HFG Workings', 'Journals', 'Component Non Cash Adjustments',
    'Other Non Cash Adjustments', 'Dvpt Non Cash Adjustments',
    'Total BSheet Adj TB', 'Total Non Cash Trial Balance',
    'Total Cash Trial Balance', 'Total Trial Balance'
)
Add-WorkingsItems $Definitions 'Output Workings' @(
    'OW - Covenant Workings', 'OW - Charts Source Data',
    'OW - Live Covenant Calculation', 'OW - Live Stress Reporting',
    'OW - Multivariable Factors', 'OW - Captured Data',
    'OW - Covenant Calculation', 'OW - MultiVar Desc Calculation',
    'FFRW - Rental Income', 'FFRW - Stock Numbers', 'FFRW - Development',
    'FFRW - Average Rents', 'FFRW - Gift Aid'
)

$AlternativeNames = @{
    'OW - Captured Data' = @(
        'StressLiveInfo', 'S0Data', 'S1Data', 'S2Data', 'S3Data', 'S4Data',
        'S5Data', 'S6Data', 'S7Data', 'S8Data', 'S9Data', 'S10Data'
    )
    'Total Non Cash Trial Balance' = @('NonCash')
    'Total Cash Trial Balance' = @('CashTB')
}

$RepositoryRoot = Split-Path -Parent $PSScriptRoot
$StructurePath = Join-Path $RepositoryRoot 'Structure.xml'
$WorkbookPath = Join-Path $RepositoryRoot 'Library\TestFileClean.xlsb'
$StructureText = [System.IO.File]::ReadAllText($StructurePath)
$ExistingGroupMatch = [regex]::Match(
    $StructureText,
    '(?s)  <GroupStructure Name="Workings">.*?^  </GroupStructure>',
    [System.Text.RegularExpressions.RegexOptions]::Multiline)
if (-not $ExistingGroupMatch.Success) { throw 'Could not locate the Workings GroupStructure.' }

$PreservedChildren = @{}
foreach ($Definition in $Definitions | Select-Object -First 10) {
    $EscapedName = [regex]::Escape((ConvertTo-XmlText $Definition.Name))
    $ChildMatch = [regex]::Match(
        $ExistingGroupMatch.Value,
        "(?s)    <ChildStructure Name=""$EscapedName"">.*?^    </ChildStructure>",
        [System.Text.RegularExpressions.RegexOptions]::Multiline)
    if (-not $ChildMatch.Success) { throw "Could not preserve existing Stock child: $($Definition.Name)" }
    $PreservedChildren[$Definition.Name] = $ChildMatch.Value + [Environment]::NewLine
}

$Excel = $null
$Workbook = $null
$GeneratedChildren = New-Object System.Collections.Generic.List[string]
$TotalSections = 0

try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.EnableEvents = $false
    $Excel.AskToUpdateLinks = $false
    $Excel.AutomationSecurity = 3
    $Workbook = $Excel.Workbooks.Open($WorkbookPath, 0, $true)

    $WorkingsNamesBySheet = @{}
    $DefinedNamesByShortName = @{}
    foreach ($DefinedName in $Workbook.Names) {
        $ShortName = Get-ShortDefinedName ([string]$DefinedName.Name)
        $DefinedNamesByShortName[$ShortName] = $DefinedName
        if ($ShortName -notlike 'Workings*') { continue }
        $WorksheetName = Get-NameWorksheet ([string]$DefinedName.RefersTo)
        if ([string]::IsNullOrWhiteSpace($WorksheetName)) { continue }
        if (-not $WorkingsNamesBySheet.ContainsKey($WorksheetName)) {
            $WorkingsNamesBySheet[$WorksheetName] = New-Object System.Collections.Generic.List[string]
        }
        $WorkingsNamesBySheet[$WorksheetName].Add($ShortName)
    }

    for ($Index = 0; $Index -lt $Definitions.Count; $Index++) {
        $Definition = $Definitions[$Index]
        if ($Index -lt 10) {
            $GeneratedChildren.Add($PreservedChildren[$Definition.Name])
            $TotalSections += ([regex]::Matches($PreservedChildren[$Definition.Name], '<CSInterfaceSection ').Count)
            continue
        }

        $Sections = New-Object System.Collections.Generic.List[string]
        $CandidateNames = @()
        if ($WorkingsNamesBySheet.ContainsKey($Definition.Worksheet)) {
            $CandidateNames = @($WorkingsNamesBySheet[$Definition.Worksheet])
            $NormalizedBase = ('Workings_' + ($Definition.Worksheet -replace '[^A-Za-z0-9]', '')).ToLowerInvariant()
            $HasVariants = @($CandidateNames | Where-Object {
                ($_ -replace '[^A-Za-z0-9]', '').ToLowerInvariant() -ne
                    ($NormalizedBase -replace '[^A-Za-z0-9]', '')
            }).Count -gt 0
            if ($HasVariants) {
                $CandidateNames = @($CandidateNames | Where-Object {
                    ($_ -replace '[^A-Za-z0-9]', '').ToLowerInvariant() -ne
                        ($NormalizedBase -replace '[^A-Za-z0-9]', '')
                })
            }
            $CandidateNames = @($CandidateNames | Sort-Object @{
                Expression = {
                    $Suffix = $_ -creplace '^.*?([A-Z]+)$', '$1'
                    Convert-LetterSuffixToNumber $Suffix
                }
            }, @{ Expression = { $_ } })
        }
        elseif ($AlternativeNames.ContainsKey($Definition.Worksheet)) {
            $CandidateNames = @($AlternativeNames[$Definition.Worksheet])
        }

        $UsedSectionNames = New-Object 'System.Collections.Generic.HashSet[string]'
        if ($CandidateNames.Count -gt 0) {
            foreach ($SourceName in $CandidateNames) {
                if (-not $DefinedNamesByShortName.ContainsKey($SourceName)) {
                    throw "Defined name '$SourceName' was not found."
                }
                $RangeInfo = Get-NameRangeInfo $Workbook $DefinedNamesByShortName[$SourceName] $Definition.Worksheet
                if ($null -eq $RangeInfo) { throw "Defined name '$SourceName' did not resolve to worksheet ranges." }
                $Fallback = if ($CandidateNames.Count -eq 1) {
                    $Definition.Name
                } else {
                    $Suffix = $SourceName -creplace '^.*?([A-Z]+)$', '$1'
                    $Definition.Name + ' ' + $Suffix
                }
                $SectionName = if ($CandidateNames.Count -eq 1) {
                    $Definition.Name
                } elseif ($SourceName -like 'Workings*') {
                    Get-BlockTitle $Workbook $Definition.Worksheet $RangeInfo $Fallback
                } else {
                    Get-AlternativeLabel $SourceName $Fallback
                }
                if (-not $UsedSectionNames.Add($SectionName)) {
                    $SectionName = $Fallback
                    $UsedSectionNames.Add($SectionName) | Out-Null
                }
                $Sections.Add((New-LiveGridSectionXml -SectionName $SectionName -Worksheet $Definition.Worksheet -SourceName $SourceName -SourceRanges '' -DataRange $RangeInfo.FirstAddress -HeaderRows ''))
            }
        }
        else {
            $DirectInfo = Get-DirectRangeInfo $Workbook $Definition.Worksheet
            $Sections.Add((New-LiveGridSectionXml -SectionName $Definition.Name -Worksheet $Definition.Worksheet -SourceName '' -SourceRanges $DirectInfo.Address -DataRange $DirectInfo.Address -HeaderRows ([string]$DirectInfo.HeaderRow)))
        }

        $GeneratedChildren.Add((New-ChildXml -Index $Index -Definition $Definition -Sections $Sections.ToArray()))
        $TotalSections += $Sections.Count
    }
}
finally {
    if ($Workbook) { $Workbook.Close($false) }
    if ($Excel) { $Excel.Quit() }
    if ($Workbook) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Workbook) | Out-Null }
    if ($Excel) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$NewGroup = @"
  <GroupStructure Name="Workings">
    <GSName>Workings</GSName>
    <GSID>1</GSID>
    <FirstChild>0</FirstChild>
$($GeneratedChildren -join '')  </GroupStructure>
"@

$NewStructureText = $StructureText.Substring(0, $ExistingGroupMatch.Index) +
    $NewGroup +
    $StructureText.Substring($ExistingGroupMatch.Index + $ExistingGroupMatch.Length)

if ($Apply) {
    [System.IO.File]::WriteAllText($StructurePath, $NewStructureText, [System.Text.UTF8Encoding]::new($false))
}

[pscustomobject]@{
    Applied = [bool]$Apply
    Children = $Definitions.Count
    Sections = $TotalSections
    RemovedDuplicates = 1
    SourceWorkbook = $WorkbookPath
    Structure = $StructurePath
} | Format-List
