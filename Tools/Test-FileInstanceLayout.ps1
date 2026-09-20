$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repo 'bin\Debug'
$output = Join-Path $repo ('obj\FileInstanceLayoutTests\' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $output)
$runner = Join-Path $output 'Runner.exe'
$references = @('System.Windows.Forms','System.Drawing','System.Core','Microsoft.CSharp')
$references += @('DevExpress.XtraBars.v25.2.dll','DevExpress.XtraEditors.v25.2.dll','DevExpress.Utils.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Drawing.v25.2.dll') | ForEach-Object { Join-Path $bin $_ }
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'FileInstanceLayoutFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $references
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner + '.config')
& $runner $bin $output
if ($LASTEXITCODE -ne 0) { throw 'File instance layout fixture failed.' }
Write-Output "Layout renders: $output"
