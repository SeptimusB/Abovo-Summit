param([string]$Configuration = 'Release', [string]$Workbook = 'C:\Repos\Abovo Summit\Library\Demo BP v26_0001.xlsb', [switch]$StressOnly, [switch]$GroupOnly)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repo ('bin\' + $Configuration)
$output = Join-Path $repo ('obj\PresentationLayoutTests\' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $output)
$runner = Join-Path $output 'Runner.exe'
$references = @('System.Windows.Forms','System.Drawing','System.Core','System.Data','System.Xml','Microsoft.CSharp')
$references += Join-Path $bin 'DevExpress.Data.Desktop.v25.2.dll'
$references += @('DevExpress.XtraBars.v25.2.dll','DevExpress.XtraEditors.v25.2.dll','DevExpress.XtraGrid.v25.2.dll','DevExpress.Utils.v25.2.dll','DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll') | ForEach-Object { Join-Path $bin $_ }
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'PresentationLayoutFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $references
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $output
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner + '.config')
$hash = (Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash
Write-Output "Layout renders: $output"
& $runner $bin $Workbook $output $(if ($StressOnly) { '--stress-only' } elseif ($GroupOnly) { '--group-only' } else { '--all' })
$runnerExit = $LASTEXITCODE
if ((Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash -ne $hash) { throw 'Source workbook changed.' }
if ($runnerExit -ne 0) { throw 'Presentation layout fixture failed; source hash remains unchanged.' }
Write-Output "PASS: source unchanged; renders: $output"
