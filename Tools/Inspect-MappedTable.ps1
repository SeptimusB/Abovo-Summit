param([string]$Workbook = 'C:\Repos\Abovo Summit\Library\Demo BP v26_0001.xlsb')
$ErrorActionPreference = 'Stop'
[void][Reflection.Assembly]::LoadFrom('C:\Repos\Abovo Summit\bin\Debug\DevExpress.Docs.v25.2.dll')
$book = New-Object DevExpress.Spreadsheet.Workbook
try {
    $book.Options.CalculationMode = [DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
    [void]$book.LoadDocument($Workbook)
    $sheet = $book.Worksheets['Check Sheet']
    Write-Output ('Used range: ' + $sheet.GetUsedRange().GetReferenceA1())
    for ($row = 0; $row -lt [Math]::Min($sheet.GetUsedRange().BottomRowIndex + 1, 120); $row++) {
        $cells = for ($column = 0; $column -lt 8; $column++) { $sheet.Cells[$row,$column].DisplayText }
        if (($cells -join '').Trim()) { Write-Output (($row + 1).ToString() + ': ' + ($cells -join ' | ')) }
    }
    foreach ($link in $sheet.Hyperlinks) {
        Write-Output ('LINK: ' + $link.Range.GetReferenceA1() + ' uri=' + $link.Uri + ' location=' + $link.Location)
    }
    foreach ($address in @('A6','A8','A9','E9','F9','G9','A63')) {
        $cell = $sheet.Cells[$address]
        Write-Output ("STYLE {0}: {1} {2}pt bold={3}, fg={4}, bg={5}" -f $address,$cell.Font.Name,$cell.Font.Size,$cell.Font.Bold,$cell.Font.Color,$cell.Fill.BackgroundColor)
    }
    foreach ($rule in $sheet.ConditionalFormattings) {
        Write-Output ("FORMAT RULE: " + $rule.GetType().Name + ' ' + $rule.Range.GetReferenceA1() + ' expression=' + $rule.Expression + ' fg=' + $rule.Formatting.Font.Color + ' bg=' + $rule.Formatting.Fill.BackgroundColor)
    }
    foreach ($rule in $sheet.DataValidations) {
        Write-Output ('VALIDATION: ' + $rule.Range.GetReferenceA1() + ' type=' + $rule.ValidationType + ' text=' + $rule.Criteria.TextValue + ' formula=' + $rule.Criteria.FormulaInvariant)
    }
    for ($row = 6; $row -lt 63; $row++) {
        for ($column = 0; $column -lt 7; $column++) {
            $cell = $sheet.Cells[$row,$column]
            $validation = $sheet.DataValidations.GetDataValidation($cell)
            if ($null -ne $validation) {
                Write-Output ('INPUT: ' + [char](65 + $column) + ($row + 1) + ' value=' + $cell.DisplayText + ' locked=' + $cell.Protection.Locked + ' formula=' + $cell.FormulaInvariant)
            }
        }
    }
} finally { $book.Dispose() }
