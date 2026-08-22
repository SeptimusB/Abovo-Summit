param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot '..\Library\TestFileMigrated.xlsb'),
    [string]$OutputDirectory = (Join-Path $env:TEMP 'AbovoFundingDashboardInspection')
)

$ErrorActionPreference = 'Stop'
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$hashBefore = (Get-FileHash -LiteralPath $resolvedWorkbook -Algorithm SHA256).Hash
[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$jsonPath = Join-Path $OutputDirectory 'FundingDashboard.json'
$pdfPath = Join-Path $OutputDirectory 'FundingDashboard.pdf'

$excel = $null
$workbook = $null
$worksheet = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    try { $excel.Calculation = -4135 } catch {}
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)
    $worksheet = $workbook.Worksheets.Item('Funding Dashboard')

    $cells = [Collections.Generic.List[object]]::new()
    foreach ($cell in $worksheet.UsedRange.Cells) {
        $formula = [string]$cell.Formula
        $value = $cell.Value2
        $displayText = [string]$cell.Text
        if ($formula -ne '' -or $null -ne $value -or $displayText -ne '') {
            $cells.Add([pscustomobject]@{
                address = $cell.Address($false, $false)
                formula = $formula
                value = $value
                displayText = $displayText
                numberFormat = [string]$cell.NumberFormat
                locked = [bool]$cell.Locked
                horizontalAlignment = $cell.HorizontalAlignment
                verticalAlignment = $cell.VerticalAlignment
                fontName = [string]$cell.Font.Name
                fontSize = $cell.Font.Size
                fontBold = [bool]$cell.Font.Bold
                fontColor = $cell.Font.Color
                fillColor = $cell.Interior.Color
                validationType = $(try { $cell.Validation.Type } catch { $null })
                validationFormula1 = $(try { [string]$cell.Validation.Formula1 } catch { '' })
            })
        }
    }

    $shapes = [Collections.Generic.List[object]]::new()
    foreach ($shape in $worksheet.Shapes) {
        $linkedCell = $null
        $listFillRange = $null
        $controlValue = $null
        $minimum = $null
        $maximum = $null
        $smallChange = $null
        $largeChange = $null
        try { $linkedCell = [string]$shape.ControlFormat.LinkedCell } catch {}
        try { $listFillRange = [string]$shape.ControlFormat.ListFillRange } catch {}
        try { $controlValue = $shape.ControlFormat.Value } catch {}
        try { $minimum = $shape.ControlFormat.Min } catch {}
        try { $maximum = $shape.ControlFormat.Max } catch {}
        try { $smallChange = $shape.ControlFormat.SmallChange } catch {}
        try { $largeChange = $shape.ControlFormat.LargeChange } catch {}
        $shapes.Add([pscustomobject]@{
            name = [string]$shape.Name
            type = $shape.Type
            formControlType = $(try { $shape.FormControlType } catch { $null })
            text = $(try { [string]$shape.TextFrame2.TextRange.Text } catch { '' })
            alternativeText = [string]$shape.AlternativeText
            onAction = [string]$shape.OnAction
            topLeftCell = [string]$shape.TopLeftCell.Address($false, $false)
            bottomRightCell = [string]$shape.BottomRightCell.Address($false, $false)
            left = $shape.Left
            top = $shape.Top
            width = $shape.Width
            height = $shape.Height
            linkedCell = $linkedCell
            listFillRange = $listFillRange
            value = $controlValue
            minimum = $minimum
            maximum = $maximum
            smallChange = $smallChange
            largeChange = $largeChange
        })
    }

    $charts = [Collections.Generic.List[object]]::new()
    foreach ($chartObject in $worksheet.ChartObjects()) {
        $series = [Collections.Generic.List[object]]::new()
        foreach ($seriesItem in $chartObject.Chart.SeriesCollection()) {
            $series.Add([pscustomobject]@{
                name = [string]$seriesItem.Name
                formula = [string]$seriesItem.Formula
                chartType = $seriesItem.ChartType
                axisGroup = $seriesItem.AxisGroup
            })
        }
        $charts.Add([pscustomobject]@{
            name = [string]$chartObject.Name
            title = $(if ($chartObject.Chart.HasTitle) { [string]$chartObject.Chart.ChartTitle.Text } else { '' })
            chartType = $chartObject.Chart.ChartType
            topLeftCell = [string]$chartObject.TopLeftCell.Address($false, $false)
            bottomRightCell = [string]$chartObject.BottomRightCell.Address($false, $false)
            left = $chartObject.Left
            top = $chartObject.Top
            width = $chartObject.Width
            height = $chartObject.Height
            series = @($series)
        })
    }

    $names = [Collections.Generic.List[object]]::new()
    foreach ($name in $workbook.Names) {
        $refersTo = [string]$name.RefersTo
        if ($refersTo -like "*'Funding Dashboard'!*" -or $refersTo -like '*Funding Dashboard!*' -or
            [string]$name.Name -in @('Funders2', 'Facility2', 'CovenantDashboard')) {
            $resolvedValues = @()
            try {
                foreach ($nameCell in $name.RefersToRange.Cells) {
                    if ([string]$nameCell.Text -ne '') { $resolvedValues += [string]$nameCell.Text }
                }
            } catch {}
            $names.Add([pscustomobject]@{
                name = [string]$name.Name
                refersTo = $refersTo
                visible = [bool]$name.Visible
                values = $resolvedValues
            })
        }
    }

    $merges = [Collections.Generic.HashSet[string]]::new()
    foreach ($cell in $worksheet.UsedRange.Cells) {
        if ([bool]$cell.MergeCells) { [void]$merges.Add([string]$cell.MergeArea.Address($false, $false)) }
    }

    $rows = for ($row = 1; $row -le $worksheet.UsedRange.Rows.Count; $row++) {
        [pscustomobject]@{ row = $row; height = $worksheet.Rows.Item($row).RowHeight; hidden = [bool]$worksheet.Rows.Item($row).Hidden }
    }
    $columns = for ($column = 1; $column -le $worksheet.UsedRange.Columns.Count; $column++) {
        [pscustomobject]@{ column = $column; width = $worksheet.Columns.Item($column).ColumnWidth; hidden = [bool]$worksheet.Columns.Item($column).Hidden }
    }

    $chartSourceCells = [Collections.Generic.List[object]]::new()
    $chartSourceSheet = $workbook.Worksheets.Item('OW - Charts Source Data')
    foreach ($cell in $chartSourceSheet.Range('AH7:AS49').Cells) {
        if ([string]$cell.Formula -ne '' -or $null -ne $cell.Value2 -or [string]$cell.Text -ne '') {
            $chartSourceCells.Add([pscustomobject]@{
                address = $cell.Address($false, $false)
                formula = [string]$cell.Formula
                value = $cell.Value2
                displayText = [string]$cell.Text
                numberFormat = [string]$cell.NumberFormat
            })
        }
    }

    $result = [pscustomobject]@{
        workbook = $resolvedWorkbook
        workbookHashBefore = $hashBefore
        sheet = [string]$worksheet.Name
        usedRange = [string]$worksheet.UsedRange.Address($false, $false)
        printArea = [string]$worksheet.PageSetup.PrintArea
        cells = @($cells)
        shapes = @($shapes)
        charts = @($charts)
        names = @($names)
        mergedRanges = @($merges)
        rows = @($rows)
        columns = @($columns)
        chartSourceCells = @($chartSourceCells)
    }
    $result | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    $worksheet.ExportAsFixedFormat(0, $pdfPath, 0, $true, $true)
}
finally {
    if ($workbook -ne $null) { $workbook.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }
    foreach ($item in @($worksheet, $workbook, $excel)) {
        if ($item -ne $null) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$hashAfter = (Get-FileHash -LiteralPath $resolvedWorkbook -Algorithm SHA256).Hash
if ($hashAfter -ne $hashBefore) { throw 'Workbook hash changed during read-only inspection.' }
Write-Output "JSON=$jsonPath"
Write-Output "PDF=$pdfPath"
Write-Output "SHA256=$hashAfter"
