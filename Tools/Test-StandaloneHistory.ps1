param([string]$Configuration = 'Release', [string]$Workbook = 'C:\Repos\Abovo Summit\Library\Demo BP v26_0001.xlsb')
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repo ('bin\' + $Configuration)
$output = Join-Path $repo ('obj\StandaloneHistoryTests\' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $output)
$runner = Join-Path $output 'Runner.exe'
$references = @('System.Windows.Forms','System.Drawing','System.Core','System.Data','System.Xml','Microsoft.CSharp')
$references += @('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll') | ForEach-Object { Join-Path $bin $_ }
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'StandaloneHistoryFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $references
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $output
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner + '.config')
$hash = (Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash
& $runner $bin $Workbook
if ($LASTEXITCODE -ne 0) { throw 'Standalone history fixture failed.' }
if ((Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash -ne $hash) { throw 'Source workbook changed.' }
Write-Output 'PASS: source workbook SHA-256 unchanged; no Save/SaveAs was used.'
