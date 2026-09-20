param([string]$Workbook = 'C:\Repos\Abovo Summit\Library\Demo BP v26_0001.xlsb', [switch]$CheckSheet)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$binaryDirectory = Join-Path $repo 'bin\Debug'
$fixtureDirectory = Join-Path $repo ('obj\MappedTableTests\' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDirectory)
$runner = Join-Path $fixtureDirectory 'Runner.exe'
$references = @('System.Windows.Forms','System.Drawing','System.Core','Microsoft.CSharp')
$references += @('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll') | ForEach-Object { Join-Path $binaryDirectory $_ }
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'MappedTableFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $references
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $fixtureDirectory
Copy-Item -LiteralPath (Join-Path $binaryDirectory 'Abovo-summit.exe.config') -Destination ($runner + '.config')
$hash = (Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash
& $runner $binaryDirectory $Workbook $(if ($CheckSheet) {'CheckSheet'} else {'Accounts'})
if ($LASTEXITCODE -ne 0) { throw 'Mapped-table fixture failed.' }
if ((Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash -ne $hash) { throw 'Source workbook changed.' }
Write-Output 'PASS: source workbook hash unchanged.'
