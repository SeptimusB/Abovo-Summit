param([Parameter(Mandatory=$true)][string]$Expanded)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo 'bin/Release'
[void][Reflection.Assembly]::LoadFrom((Join-Path $bin 'Abovo-summit.exe'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $bin 'DevExpress.Docs.v25.2.dll'))
$out=Join-Path $repo ('obj/JournalExcelRoundtrip/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$result=Join-Path $out 'excel-roundtrip.xlsb'
$hash=(Get-FileHash -LiteralPath $Expanded).Hash
$before=New-Object DevExpress.Spreadsheet.Workbook
$after=New-Object DevExpress.Spreadsheet.Workbook
$excel=$null;$seed=$null;$book=$null
try {
    $before.Options.CalculationMode=[DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
    [void]$before.LoadDocument($Expanded)
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    $excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
    $seed=$excel.Workbooks.Add();$excel.Calculation=-4135
    $book=$excel.Workbooks.Open($Expanded,0,$true)
    foreach($name in @('IR_Journals','Rep_Jour_01','TransCopy_IR_Journals_01','TransCopy_IR_Journals_02')) {
        $range=$book.Names.Item($name).RefersToRange
        if($range.Rows.Count -ne $before.DefinedNames.GetDefinedName($name).Range.RowCount){throw "Excel name differs: $name"}
        Write-Output "PASS: Excel $name rows=$($range.Rows.Count)"
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range)
    }
    $sheet=$book.Worksheets.Item('Check Sheet');$sheet.Calculate()
    foreach($address in @('B53','C53','D53','E53','F53')) {
        $cell=$sheet.Range($address)
        Write-Output "CHECK $address formula=$($cell.Formula) value=$($cell.Text)"
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
    $book.SaveCopyAs($result);$book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book);$book=$null
    $after.Options.CalculationMode=[DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
    [void]$after.LoadDocument($result)
    foreach($name in @('IR_Journals','Rep_Jour_01','TransCopy_IR_Journals_01','TransCopy_IR_Journals_02')) {
        if($before.DefinedNames.GetDefinedName($name).RefersTo -ne $after.DefinedNames.GetDefinedName($name).RefersTo){throw "Round-trip range differs: $name"}
    }
    Write-Output "PASS: Summit -> Excel read-only/macro-disabled SaveCopyAs -> native XLSB reopen. Output: $result"
} finally {
    if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
    if($seed){$seed.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($seed)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
    $before.Dispose();$after.Dispose()
    if((Get-FileHash -LiteralPath $Expanded).Hash -ne $hash){throw 'Inspection source changed'}
}
