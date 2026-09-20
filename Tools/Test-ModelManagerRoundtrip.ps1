param([string]$Repository = 'C:\Repos\Abovo Summit', [string]$CopyDirectory = 'C:\Sandbox\Model Manager Trial', [switch]$FormOnly)
$ErrorActionPreference = 'Stop'
$bin = Join-Path $Repository 'bin\Debug'
[void][Reflection.Assembly]::LoadFrom((Join-Path $bin 'Abovo-summit.exe'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $bin 'DevExpress.Docs.v25.2.dll'))
$testDirectory = Join-Path $Repository ('obj\ModelManagerTests\roundtrip-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $testDirectory)
$form = New-Object ModelManagerForm
try {
    $form.ShowInTaskbar = $false
    $form.StartPosition = [Windows.Forms.FormStartPosition]::Manual
    $form.Location = [Drawing.Point]::new(-20000, -20000)
    $form.Show()
    [Windows.Forms.Application]::DoEvents()
    $form.CreateControl()
    $form.PerformLayout()
    $bitmap = New-Object Drawing.Bitmap $form.Width, $form.Height
    try {
        $form.DrawToBitmap($bitmap, [Drawing.Rectangle]::new(0,0,$form.Width,$form.Height))
        $bitmap.Save((Join-Path $testDirectory 'manager-form.png'))
    } finally { $bitmap.Dispose() }
    Write-Output 'PASS: Native Model Manager form constructed and rendered without errors'
} finally { $form.Dispose() }
if ($FormOnly) { Write-Output ('Form render: ' + $testDirectory); exit 0 }

$excel = $null
$seed = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $excel.AskToUpdateLinks = $false
    $seed = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    foreach ($fileName in @('Stori - Model Manager Trial.xlsb', 'Blank BP v26_0001 - Model Manager Trial.xlsb')) {
        $source = Join-Path $CopyDirectory $fileName
        $before = [Abovo.ModelManagerStore]::Fingerprint($source)
        $expected = [Abovo.ModelManagerStore]::ReadEmbedded($source)
        $nativePath = Join-Path $testDirectory ('native-' + $fileName)
        $excelPath = Join-Path $testDirectory ('excel-' + $fileName)
        $book = New-Object DevExpress.Spreadsheet.Workbook
        try {
            $book.Options.CalculationMode = [DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
            [void]$book.LoadDocument($source)
            if ($book.CustomXmlParts.Count -ne 4) { throw 'Expected four XML parts (three existing and one manager).' }
            $book.SaveDocument($nativePath, [DevExpress.Spreadsheet.DocumentFormat]::Xlsb)
        } finally { $book.Dispose() }
        $excelBook = $null
        try {
            $excelBook = $excel.Workbooks.Open($nativePath, 0, $true)
            $parts = $excelBook.CustomXMLParts.SelectByNamespace([Abovo.ModelManagerStore]::XmlNamespace)
            try { if ($parts.Count -ne 1) { throw 'Excel did not find exactly one manager XML part.' } }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($parts) }
            $excelBook.SaveCopyAs($excelPath)
        } finally {
            if ($excelBook) { $excelBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelBook) }
        }
        $actual = [Abovo.ModelManagerStore]::ReadEmbedded($excelPath)
        if ($actual.DefinitionId -ne $expected.DefinitionId -or $actual.IsGeneric -ne $expected.IsGeneric -or $actual.Rules.Count -ne $expected.Rules.Count) { throw 'Definition changed during roundtrip.' }
        $book = New-Object DevExpress.Spreadsheet.Workbook
        try {
            $book.Options.CalculationMode = [DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
            [void]$book.LoadDocument($excelPath)
            if ($book.CustomXmlParts.Count -ne 4) { throw 'XML parts lost after Excel save.' }
        } finally { $book.Dispose() }
        if ([Abovo.ModelManagerStore]::Fingerprint($source) -ne $before) { throw 'Managed test input modified.' }
        Write-Output ('PASS: DevExpress -> XLSB -> Excel -> XLSB -> DevExpress, four XML parts: ' + $fileName)
    }
} finally {
    if ($seed) { $seed.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($seed) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
Write-Output ('Roundtrip artifacts: ' + $testDirectory)
