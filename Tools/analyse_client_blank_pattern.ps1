param(
    [Parameter(Mandatory = $true)][string]$MasterPath,
    [Parameter(Mandatory = $true)][string]$BlankPath
)

$ErrorActionPreference = 'Stop'
$excel = $null
$master = $null
$blank = $null
function Release-ComObject([object]$Object) { if ($null -ne $Object) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } }
function Get-ContentValue([object]$Content, [int]$Row, [int]$Column) {
    if ($Content -is [System.Array]) { return [string]$Content[$Row, $Column] }
    return [string]$Content
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $master = $excel.Workbooks.Open($MasterPath, 0, $true)
    $blank = $excel.Workbooks.Open($BlankPath, 0, $true)

    $total = 0L
    for ($index = 1; $index -le $master.Worksheets.Count; $index++) {
        $masterSheet = $null
        $blankSheet = $null
        $masterUsed = $null
        $blankRange = $null
        try {
            $masterSheet = $master.Worksheets.Item($index)
            try { $blankSheet = $blank.Worksheets.Item($masterSheet.Name) } catch { continue }
            $masterUsed = $masterSheet.UsedRange
            $blankRange = $blankSheet.Range($masterUsed.Address())
            $masterContent = $masterUsed.Formula
            $blankContent = $blankRange.Formula
            $rowLower = if ($masterContent -is [System.Array]) { $masterContent.GetLowerBound(0) } else { 1 }
            $rowUpper = if ($masterContent -is [System.Array]) { $masterContent.GetUpperBound(0) } else { 1 }
            $columnLower = if ($masterContent -is [System.Array]) { $masterContent.GetLowerBound(1) } else { 1 }
            $columnUpper = if ($masterContent -is [System.Array]) { $masterContent.GetUpperBound(1) } else { 1 }

            $count = 0L
            $locked = 0L
            $unlocked = 0L
            $samples = New-Object System.Collections.Generic.List[string]
            for ($row = $rowLower; $row -le $rowUpper; $row++) {
                for ($column = $columnLower; $column -le $columnUpper; $column++) {
                    $masterValue = Get-ContentValue $masterContent $row $column
                    $blankValue = Get-ContentValue $blankContent $row $column
                    if ($masterValue -ne '' -and -not $masterValue.StartsWith('=') -and $blankValue -eq '') {
                        $count++
                        $sheetRow = $masterUsed.Row + $row - $rowLower
                        $sheetColumn = $masterUsed.Column + $column - $columnLower
                        $cell = $masterSheet.Cells.Item($sheetRow, $sheetColumn)
                        if ([bool]$cell.Locked) { $locked++ } else { $unlocked++ }
                        if ($samples.Count -lt 8) {
                            $display = ([string]$cell.Text).Replace("`r", ' ').Replace("`n", ' ')
                            if ($display.Length -gt 50) { $display = $display.Substring(0, 50) }
                            $samples.Add("$($cell.Address($false, $false))=$display")
                        }
                        Release-ComObject $cell
                    }
                }
            }
            if ($count -gt 0) {
                $total += $count
                [pscustomobject]@{
                    Sheet = $masterSheet.Name
                    ClearCount = $count
                    Unlocked = $unlocked
                    Locked = $locked
                    Samples = ($samples -join '; ')
                }
            }
            if (($index % 50) -eq 0) { Write-Progress -Activity 'Analysing blank pattern' -Status "$index of $($master.Worksheets.Count) sheets" -PercentComplete (($index / $master.Worksheets.Count) * 100) }
        }
        finally {
            Release-ComObject $blankRange
            Release-ComObject $masterUsed
            Release-ComObject $blankSheet
            Release-ComObject $masterSheet
        }
    }
    [pscustomobject]@{ Sheet = 'TOTAL'; ClearCount = $total; Unlocked = ''; Locked = ''; Samples = '' }
}
finally {
    if ($null -ne $blank) { [void]$blank.Close($false); Release-ComObject $blank }
    if ($null -ne $master) { [void]$master.Close($false); Release-ComObject $master }
    if ($null -ne $excel) { [void]$excel.Quit(); Release-ComObject $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
