param([string]$Configuration='Release', [string]$Workbook=(Join-Path (Split-Path -Parent $PSScriptRoot) 'Library/Demo BP v26_0001.xlsb'))
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/AnalyserChartTests/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$runner=Join-Path $out 'Runner.exe'
$refs=@('System.Windows.Forms','System.Drawing','System.Core','System.Data','System.Xml','Microsoft.CSharp')
$refs+=@('DevExpress.Data.Desktop.v25.2.dll','DevExpress.XtraGrid.v25.2.dll','DevExpress.XtraCharts.v25.2.dll','DevExpress.XtraCharts.v25.2.UI.dll','DevExpress.XtraEditors.v25.2.dll','DevExpress.Utils.v25.2.dll','DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll') | ForEach-Object { Join-Path $bin $_ }
$refs += Join-Path $bin 'DevExpress.Charts.v25.2.Core.dll'
$refs += Join-Path $bin 'DevExpress.XtraBars.v25.2.dll'
$refs += Join-Path $bin 'DevExpress.XtraTreeList.v25.2.dll'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'AnalyserChartFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
$hash=(Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash
Write-Output ('Output: '+$out)
try {
    & $runner $bin $Workbook $out
    if($LASTEXITCODE -ne 0){throw 'Analyser chart fixture failed'}
} finally {
    if((Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash -ne $hash){throw 'Source workbook changed'}
}
Write-Output 'PASS: source workbook hash unchanged.'
