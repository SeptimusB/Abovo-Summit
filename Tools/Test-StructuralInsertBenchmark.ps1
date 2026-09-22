param([string]$Configuration='Release')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/StructuralInsertBenchmarkTests/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$runner=Join-Path $out 'Runner.exe'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'StructuralInsertBenchmarkFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies @('System.Core')
& $runner $bin
if($LASTEXITCODE -ne 0){throw 'Structural insertion benchmark fixture failed'}
