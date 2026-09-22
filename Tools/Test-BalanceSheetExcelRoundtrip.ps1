param([Parameter(Mandatory=$true)][string]$Snapshot, [string]$Configuration='Release')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
[void][Reflection.Assembly]::LoadFrom((Join-Path $bin 'Abovo-summit.exe'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $bin 'DevExpress.Docs.v25.2.dll'))
Add-Type -AssemblyName System.IO.Compression.FileSystem
$out=Join-Path $repo ('obj/BalanceSheetExcelRoundtrip/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$result=Join-Path $out 'excel-roundtrip.xlsb'
$hash=(Get-FileHash -LiteralPath $Snapshot).Hash
function Read-PackageEntryHash([string]$Path,[string]$Name) {
    $zip=[IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry=$zip.GetEntry($Name)
        if(!$entry){return ''}
        $stream=$entry.Open()
        $sha=[Security.Cryptography.SHA256]::Create()
        try { return [BitConverter]::ToString($sha.ComputeHash($stream)) }
        finally {$stream.Dispose();$sha.Dispose()}
    } finally {$zip.Dispose()}
}
$original=New-Object DevExpress.Spreadsheet.Workbook
$reopened=New-Object DevExpress.Spreadsheet.Workbook
$excel=$null;$seed=$null;$excelBook=$null
try {
    $original.Options.CalculationMode=[DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
    [void]$original.LoadDocument($Snapshot)
    $before=[Abovo.BalanceSheetSnapshot]::Read($original,[Abovo.BalanceSheetStatement]::Read($original))
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    $excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
    $seed=$excel.Workbooks.Add();$excel.Calculation=-4135
    $excelBook=$excel.Workbooks.Open($Snapshot,0,$true)
    $excelBook.SaveCopyAs($result)
    $excelBook.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelBook);$excelBook=$null
    $reopened.Options.CalculationMode=[DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
    [void]$reopened.LoadDocument($result)
    $after=[Abovo.BalanceSheetSnapshot]::Read($reopened,[Abovo.BalanceSheetStatement]::Read($reopened))
    if($before.Nodes.Count -ne $after.Nodes.Count){throw 'Snapshot hierarchy changed'}
    for($i=0;$i -lt $before.Nodes.Count;$i++){
        if($before.Nodes[$i].Id -ne $after.Nodes[$i].Id){throw 'Snapshot identity changed'}
        for($p=0;$p -lt 41;$p++){if($before.Nodes[$i].Values[$p] -ne $after.Nodes[$i].Values[$p]){throw 'Frozen value changed'}}
    }
    if($original.Worksheets.Count -ne $reopened.Worksheets.Count){throw 'Worksheet count changed'}
    for($i=0;$i -lt $original.Worksheets.Count;$i++){
        if($original.Worksheets[$i].Name -ne $reopened.Worksheets[$i].Name){throw 'Worksheet order changed'}
    }
    foreach($name in $original.DefinedNames){
        $actual=$reopened.DefinedNames.GetDefinedName($name.Name)
        if(!$actual -or $actual.RefersTo -ne $name.RefersTo){throw ('Named range changed: '+$name.Name)}
    }
    if((Read-PackageEntryHash $Snapshot 'xl/vbaProject.bin') -ne (Read-PackageEntryHash $result 'xl/vbaProject.bin')){
        Write-Output 'NOTE: Excel rewrote the VBA binary container. Verify module source hashes with Verify-VbaModuleHashes.py; binary inequality alone does not prove a source change.'
    }
    Write-Output 'PASS: native Excel read-only open/SaveCopyAs/Summit reload; complete frozen hierarchy/values, sheet order and global names preserved.'
    Write-Output 'Excel macros/events and calculation were disabled. Interactive Excel/VBA calculation remains a manual acceptance test.'
    Write-Output ('Artifact: '+$result)
} finally {
    if($excelBook){$excelBook.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelBook)}
    if($seed){$seed.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($seed)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
    $original.Dispose();$reopened.Dispose()
    if((Get-FileHash -LiteralPath $Snapshot).Hash -ne $hash){throw 'Input snapshot modified'}
}
