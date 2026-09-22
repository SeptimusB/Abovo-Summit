param([string]$Configuration='Release', [string]$Workbook=(Join-Path (Split-Path -Parent $PSScriptRoot) 'Library/Demo BP v26_0001.xlsb'), [switch]$SaveOnly)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/EditorNavigationTests/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$runner=Join-Path $out 'Runner.exe'
$refs=@('System.Windows.Forms','System.Drawing','System.Core','System.Data','System.Xml','Microsoft.CSharp')
$refs+=Join-Path $bin 'DevExpress.XtraBars.v25.2.dll'
$refs+=Join-Path $bin 'DevExpress.Docs.v25.2.dll'
$refs+=@('DevExpress.Data.Desktop.v25.2.dll','DevExpress.XtraGrid.v25.2.dll','DevExpress.XtraVerticalGrid.v25.2.dll','DevExpress.XtraEditors.v25.2.dll','DevExpress.Utils.v25.2.dll','DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll') | ForEach-Object { Join-Path $bin $_ }
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'EditorNavigationFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
$hash=(Get-FileHash -LiteralPath $Workbook).Hash
Write-Output ('Output: '+$out)
try {
    $testMode=if($SaveOnly){'save-only'}else{'full'}
    & $runner $bin $Workbook $out $testMode
    if($LASTEXITCODE -ne 0){throw 'Editor navigation fixture failed'}
} finally {
    if((Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Source workbook changed'}
    Write-Output 'Source hash unchanged'
}
