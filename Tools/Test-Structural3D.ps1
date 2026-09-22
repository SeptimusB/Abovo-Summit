param([string]$Configuration='Release')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/Structural3DTests/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Core','System.Drawing')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'Structural3DFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
& $runner $bin
if($LASTEXITCODE -ne 0){throw '3-D fixture failed'}
